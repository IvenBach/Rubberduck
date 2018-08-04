﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Rubberduck.Common;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Resources.Registration;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using Rubberduck.Root;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.API.VBA
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.IParserGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IParser
    {
        [DispId(1)]
        void Parse();
        [DispId(2)]
        void BeginParse();
        [DispId(3)]
        Declarations AllDeclarations { get; }
        [DispId(4)]
        Declarations UserDeclarations { get; }
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.IParserEventsGuid),
        InterfaceType(ComInterfaceType.InterfaceIsIDispatch)
    ]
    public interface IParserEvents
    {
        [DispId(1)]
        void OnStateChanged(ParserState CurrentState);
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.ParserClassGuid),
        ProgId(RubberduckProgId.ParserStateProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IParser)),
        ComSourceInterfaces(typeof(IParserEvents)),
        EditorBrowsable(EditorBrowsableState.Always)
    ]
    public sealed class Parser : IParser, IDisposable
    {
        private RubberduckParserState _state;
        private SynchronousParseCoordinator _parser;
        private IVBE _vbe;
        private IVBEEvents _vbeEvents;
        private readonly IUiDispatcher _dispatcher;
        private readonly CancellationTokenSource _tokenSource;

        internal Parser()
        {
            UiContextProvider.Initialize();
            _dispatcher = new UiDispatcher(UiContextProvider.Instance());
            _tokenSource = new CancellationTokenSource();
        }

        // vbe is the com coclass interface from the interop assembly.
        // There is no shared interface between VBA and VB6 types, hence object.
        internal Parser(object vbe) : this()
        {
            if (_parser != null)
            {
                throw new InvalidOperationException("ParserState is already initialized.");
            }

            _vbe = RootComWrapperFactory.GetVbeWrapper(vbe);
            _vbeEvents = VBEEvents.Initialize(_vbe);
            var declarationFinderFactory = new ConcurrentlyConstructedDeclarationFinderFactory();
            var projectRepository = new ProjectsRepository(_vbe);
            _state = new RubberduckParserState(_vbe, projectRepository, declarationFinderFactory, _vbeEvents);
            _state.StateChanged += _state_StateChanged;

            var sourceFileHandler = _vbe.SourceFileHandler;
            var compilationArgumentsProvider = new CompilationArgumentsProvider(projectRepository, _dispatcher);
            var compilationsArgumentsCache = new CompilationArgumentsCache(compilationArgumentsProvider);
            var preprocessorErrorListenerFactory = new PreprocessingParseErrorListenerFactory();
            var preprocessorParser = new VBAPreprocessorParser(preprocessorErrorListenerFactory, preprocessorErrorListenerFactory);
            var preprocessor = new VBAPreprocessor(double.Parse(_vbe.Version, CultureInfo.InvariantCulture), preprocessorParser, compilationsArgumentsCache);
            var mainParseErrorListenerFactory = new MainParseErrorListenerFactory();
            var mainTokenStreamParser = new VBATokenStreamParser(mainParseErrorListenerFactory, mainParseErrorListenerFactory);
            var tokenStreamProvider = new SimpleVBAModuleTokenStreamProvider();
            var parser = new TokenStreamParserStringParserAdapterWithPreprocessing(tokenStreamProvider, mainTokenStreamParser, preprocessor);
            var attributesSourceCodeHandler = new SourceFileHandlerSourceCodeHandlerAdapter(sourceFileHandler, projectRepository);
            var projectManager = new RepositoryProjectManager(projectRepository);
            var moduleToModuleReferenceManager = new ModuleToModuleReferenceManager();
            var parserStateManager = new ParserStateManager(_state);
            var referenceRemover = new ReferenceRemover(_state, moduleToModuleReferenceManager);
            var supertypeClearer = new SupertypeClearer(_state);
            var comSynchronizer = new COMReferenceSynchronizer(_state, parserStateManager);
            var builtInDeclarationLoader = new BuiltInDeclarationLoader(
                _state,
                new List<ICustomDeclarationLoader>
                    {
                        new DebugDeclarations(_state),
                        new SpecialFormDeclarations(_state),
                        new FormEventDeclarations(_state),
                        new AliasDeclarations(_state),
                        //new RubberduckApiDeclarations(_state)
                    }
                );
            var codePaneSourceCodeHandler = new CodePaneSourceCodeHandler(projectRepository);
            var moduleRewriterFactory = new ModuleRewriterFactory(
                codePaneSourceCodeHandler,
                attributesSourceCodeHandler);
            var parseRunner = new ParseRunner(
                _state,
                parserStateManager,
                parser,
                codePaneSourceCodeHandler,
                attributesSourceCodeHandler,
                moduleRewriterFactory);
            var declarationResolveRunner = new DeclarationResolveRunner(
                _state, 
                parserStateManager, 
                comSynchronizer);
            var referenceResolveRunner = new ReferenceResolveRunner(
                _state,
                parserStateManager,
                moduleToModuleReferenceManager,
                referenceRemover);
            var parsingStageService = new ParsingStageService(
                comSynchronizer,
                builtInDeclarationLoader,
                parseRunner,
                declarationResolveRunner,
                referenceResolveRunner  
                );
            var parsingCacheService = new ParsingCacheService(
                _state,
                moduleToModuleReferenceManager,
                referenceRemover,
                supertypeClearer,
                compilationsArgumentsCache
                );

            _parser = new SynchronousParseCoordinator(
                _state,
                parsingStageService,
                parsingCacheService,
                projectManager,
                parserStateManager
                );
        }
        
        /// <summary>
        /// Blocking call, for easier unit-test code
        /// </summary>
        public void Parse()
        {
            _parser.Parse(_tokenSource);
        }

        /// <summary>
        /// Begins asynchronous parsing
        /// </summary>
        public void BeginParse()
        {
            // non-blocking call
            _dispatcher.Invoke(() => _state.OnParseRequested(this));
        }
        
        public delegate void OnStateChangedDelegate(ParserState ParserState);
        public event OnStateChangedDelegate OnStateChanged;
        
        private void _state_StateChanged(object sender, EventArgs e)
        {
            AllDeclarations = new Declarations(_state.AllDeclarations
                .Select(item => new Declaration(item)));

            UserDeclarations = new Declarations(_state.AllUserDeclarations
                .Select(item => new Declaration(item)));

            var state = (ParserState) _state.Status;
            var stateHandler = OnStateChanged;
            if (stateHandler != null)
            {
                _dispatcher.RaiseComEvent(() =>
                {
                    stateHandler.Invoke(state);
                });
            }
        }

        public Declarations AllDeclarations { get; private set; }

        public Declarations UserDeclarations { get; private set; }

        private bool _disposed;
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            if (_state != null)
            {
                _state.StateChanged -= _state_StateChanged;
            }
            
            _disposed = true;
        }
    }
}

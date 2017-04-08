﻿using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.VBA
{
    public abstract class ParserStateManagerBase:IParserStateManager 
    {
        protected readonly RubberduckParserState _state;

        public ParserStateManagerBase(RubberduckParserState state)
        {
            if (state == null) throw new ArgumentException(nameof(state));

            _state = state;
        }


        public abstract void SetModuleStates(ICollection<QualifiedModuleName> modules, ParserState parserState, CancellationToken token, bool evaluateOverallParserState = true);


        public ParserState OverallParserState
        {
            get
            {
                return _state.Status;
            }
        }

        public void EvaluateOverallParserState(CancellationToken token)
        {
            _state.EvaluateParserState();
        }

        public void SetModuleState(QualifiedModuleName module, ParserState parserState, CancellationToken token, bool evaluateOverallParserState = true)
        {
            _state.SetModuleState(module.Component, parserState, token, null, evaluateOverallParserState);
        }

        public void SetStatusAndFireStateChanged(object requestor, ParserState status, CancellationToken token)
        {
            _state.SetStatusAndFireStateChanged(requestor, status);
        }
    }
}

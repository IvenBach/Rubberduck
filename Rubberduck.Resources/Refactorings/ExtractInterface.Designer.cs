﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Rubberduck.Resources.Refactorings {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "16.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    public class ExtractInterface {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal ExtractInterface() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("Rubberduck.Resources.Refactorings.ExtractInterface", typeof(ExtractInterface).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Rubberduck - Extract Interface.
        /// </summary>
        public static string Caption {
            get {
                return ResourceManager.GetString("Caption", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Instancing.
        /// </summary>
        public static string InstancingGroupBox {
            get {
                return ResourceManager.GetString("InstancingGroupBox", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Please specify interface name and members..
        /// </summary>
        public static string InstructionLabel {
            get {
                return ResourceManager.GetString("InstructionLabel", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Members.
        /// </summary>
        public static string MembersGroupBox {
            get {
                return ResourceManager.GetString("MembersGroupBox", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Private.
        /// </summary>
        public static string Private {
            get {
                return ResourceManager.GetString("Private", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Public.
        /// </summary>
        public static string Public {
            get {
                return ResourceManager.GetString("Public", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The implementing class is &apos;Public&apos; mandating the interface be public as well.
        ///If you require a &apos;Private&apos; interface, change the classes instancing to private as well.
        ///A private class can still implement a public interface..
        /// </summary>
        public static string PublicInstancingMandatedByPublicClass {
            get {
                return ResourceManager.GetString("PublicInstancingMandatedByPublicClass", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Extract Interface.
        /// </summary>
        public static string TitleLabel {
            get {
                return ResourceManager.GetString("TitleLabel", resourceCulture);
            }
        }
    }
}
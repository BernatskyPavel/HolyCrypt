﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace StegoLine.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "17.3.0.0")]
    internal sealed partial class General : global::System.Configuration.ApplicationSettingsBase {
        
        private static General defaultInstance = ((General)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new General())));
        
        public static General Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("cx2")]
        public string NamespacePrefix {
            get {
                return ((string)(this["NamespacePrefix"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("http://schemas.microsoft.com/office/drawing/2017/4/8/chartex")]
        public string NamespaceUri {
            get {
                return ((string)(this["NamespaceUri"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("hshId")]
        public string HashAttributeName {
            get {
                return ((string)(this["HashAttributeName"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("rv")]
        public string RunAttributeName {
            get {
                return ((string)(this["RunAttributeName"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("635")]
        public int MinOutlineWidth {
            get {
                return ((int)(this["MinOutlineWidth"]));
            }
            set {
                this["MinOutlineWidth"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("1270")]
        public int MaxOutlineWidth {
            get {
                return ((int)(this["MaxOutlineWidth"]));
            }
            set {
                this["MaxOutlineWidth"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("96000")]
        public int MinOutlineAlpha {
            get {
                return ((int)(this["MinOutlineAlpha"]));
            }
            set {
                this["MinOutlineAlpha"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("100000")]
        public int MaxOutlineAlpha {
            get {
                return ((int)(this["MaxOutlineAlpha"]));
            }
            set {
                this["MaxOutlineAlpha"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("800")]
        public int WindowWidth {
            get {
                return ((int)(this["WindowWidth"]));
            }
            set {
                this["WindowWidth"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("600")]
        public int WindowHeight {
            get {
                return ((int)(this["WindowHeight"]));
            }
            set {
                this["WindowHeight"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("ru-RU")]
        public string Language {
            get {
                return ((string)(this["Language"]));
            }
            set {
                this["Language"] = value;
            }
        }
    }
}

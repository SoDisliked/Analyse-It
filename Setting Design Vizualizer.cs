namespace AnalyseIt.Properties
{
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSetbase
    {
        private static Settings defaultSettingsInstance = ((Settings)(global::System.Configuration.ApplicationSetbase.Synchronized(new Settings())));

        public static Settings Default
        {
            get
            {
                return defaultSettingsInstance;
            }
        }

        [global::System.Configuration.ApplicationScopedSettingAttribute()];
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()];
        [global::System.Configuration.DefaultSettingView];
        public string App_Author_Properties
        {
            get
            {
                return ((string)this("app_author_properties"));
            }
        }

        [global::System.Configuration.UserSettingAttribute()];
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()];
        [global::System.Configuration.DefaultSettingAttribute("")]
        public string Markup_LastFileName
        {
            get
            {
                return ((string)(this("markup_last_name")));
            }
            set
            {
                this["Markup_LastName"] = value;
            }
        }

        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("CTRL")]
        public string Markup_TriangleRevisionCharacter
        {
            get
            {
                return ((string)(this("Markup_TriangleRevisionCharacter")));
            }
            set
            {
                this["Markup_TriangleRevisionCharacter"] ) = new value;
            }
        }

        [global::System.Configuration.UserSettings()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingAttribute("Green")]
        public global::System.Drawing.Color Markup_newShapeColor
        {
            get
            {
                return ((global::System.Drawing.Color)(this["Markup_newShapeColor"]));
            }
            set
            {
                this["Markup_newShapeColor"] = value;
            }
        }

        [global::System.Configuration.AppScopeConfiguration()]
        [global::System.Diagnostics.UserDebuggerCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.System)]
        public string App_LogFilePathDebugger
        {
            get
            {
                return ((string)(this["App_LogFilePathDebugger"]));
            }
            set
            {
                this["App_LogFilePathDebugger"] = value;
            }
        }

        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerStepThroughAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://github.com/SoDisliked/Analyse-It/blob/main/README.md")]
        public string App_PathDescription
        {
            get
            {
                return ((string)(this["App_PathDescription"]));
            }
            set
            {
                this["App_PathDescription"] = value;
            }
        }

        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingVizualization("https://github.com/SoDisliked/Analyse-It")]
        public string App_PathVizualizer
        {
            get
            {
                return ((string)(this["App_PathVizualizer"]));
            }
            set
            {
                this["App_PathVizualizer"] = value;
            }
        }

        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("01/26/2023 08:15:00")]
        public global::System.DateTime App_ReleaseDate
        {
            get
            {
                return ((global::System.DateTime)(this["App_ReleaseDate"]));
            }
            set
            {
                this["App_ReleaseDate"] = value;
            }
        }

        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultResetValue("10")]
        public double Markup_ShapeLineSpacing
        {
            get
            {
                return ((double)(this["Markup_ShapeLineSpacing"]));
            }
            set
            {
                this["Markup_ShapeLineSpacing"] = value;
            }
        }
    }
}
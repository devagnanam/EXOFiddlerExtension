using Fiddler;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EXOFiddlerInspector.Services
{
    public static class Preferences
    {
        internal static List<string> Developers = new List<string>(new string[] { "jeknight", "brandev", "bever", "jasonsla", "nick", "jeremy" });

        public static void Initialize()
        {
            SetDefaultPreferences();
        }

        public static bool IsDeveloper()
        {
            return Developers.Any(Environment.UserName.Contains);
        }

        /// <summary>
        /// Return DeveloperDemoMode value.
        /// </summary>
        /// <returns>DeveloperDemoMode</returns>
        public static bool GetDeveloperMode()
        {
            bool isdevMode = Developers.Any(Environment.UserName.Contains) && System.Diagnostics.Debugger.IsAttached;

            return isdevMode;
        }

        /// <summary>
        /// This is the low water mark for what is considered a slow running session, considering a number of factors.
        /// Exchange response times are typically going to be much quicker than this. In the < 300ms range.
        /// </summary>
        public static int GetSlowRunningSessionThreshold()
        {
            return 5000;
        }

        public static void SetDefaultPreferences()
        {
            FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.ElapsedTimeColumnEnabled", true);
            FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.enabled", true);
            FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.ResponseServerColumnEnabled", true);
            FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.ExchangeTypeColumnEnabled", true);
            FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.HostIPColumnEnabled", true);
            FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.AuthColumnEnabled", true);
            FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.AppLoggingEnabled", true);
            FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.HighlightOutlookOWAOnlyEnabled", true);
            FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.LoadSaz", true);
            FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.ColumnsEnableAll", true);
            FiddlerApplication.Prefs.SetInt32Pref("extensions.EXOFiddlerExtension.ExecutionCount", 0);

        }

        public static string AppVersion
        {
           get
            {
                Assembly assembly = Assembly.GetExecutingAssembly();

                FileVersionInfo fileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);
                return fileVersionInfo.FileVersion;
            }
        }

        private static bool _extensionEnabled;
        public static bool ExtensionEnabled
        {
            get => _extensionEnabled = FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.enabled", _extensionEnabled);
            set { _extensionEnabled = value; FiddlerApplication.Prefs.SetBoolPref("extensions.EXOFiddlerExtension.enabled", value); }
        }

        private static bool _elapsedTimeColumnEnabled;
        public static bool ElapsedTimeColumnEnabled
        {
            get => _elapsedTimeColumnEnabled = FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.ElapsedTimeColumnEnabled", _elapsedTimeColumnEnabled);
            set { _elapsedTimeColumnEnabled = value; FiddlerApplication.Prefs.SetBoolPref("extensions.EXOFiddlerExtension.ElapsedTimeColumnEnabled", value); }
        }

        private static bool _responseServerColumnEnabled;
        public static bool ResponseServerColumnEnabled
        {
            get => _responseServerColumnEnabled = FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.enabled", _responseServerColumnEnabled);
            set { _responseServerColumnEnabled = value; FiddlerApplication.Prefs.SetBoolPref("extensions.EXOFiddlerExtension.enabled", value); }
        }

        private static bool _exchangeTypeColumnEnabled;
        public static bool ExchangeTypeColumnEnabled
        {
            get => _exchangeTypeColumnEnabled = FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.enabled", _exchangeTypeColumnEnabled);
            set { _exchangeTypeColumnEnabled = value; FiddlerApplication.Prefs.SetBoolPref("extensions.EXOFiddlerExtension.enabled", value); }
        }

        private static bool _hostIPColumnEnabled;
        public static bool HostIPColumnEnabled
        {
            get => _hostIPColumnEnabled = FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.enabled", _hostIPColumnEnabled);
            set { _hostIPColumnEnabled = value; FiddlerApplication.Prefs.SetBoolPref("extensions.EXOFiddlerExtension.enabled", value); }
        }

        private static bool _authColumnEnabled;
        public static bool AuthColumnEnabled
        {
            get => _authColumnEnabled = FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.enabled", _authColumnEnabled);
            set { _authColumnEnabled = value; FiddlerApplication.Prefs.SetBoolPref("extensions.EXOFiddlerExtension.enabled", value); }
        }

        private static bool _appLoggingEnabled;
        public static bool AppLoggingEnabled
        {
            get => _appLoggingEnabled = FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.enabled", _appLoggingEnabled);
            set { _appLoggingEnabled = value; FiddlerApplication.Prefs.SetBoolPref("extensions.EXOFiddlerExtension.enabled", value); }
        }

        private static bool _highlightOutlookOWAOnlyEnabled;
        public static bool HighlightOutlookOWAOnlyEnabled
        {
            get => _highlightOutlookOWAOnlyEnabled = FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.enabled", _highlightOutlookOWAOnlyEnabled);
            set { _highlightOutlookOWAOnlyEnabled = value; FiddlerApplication.Prefs.SetBoolPref("extensions.EXOFiddlerExtension.enabled", value); }
        }


        private static bool _isLoadSaz;
        public static bool IsLoadSaz
        {
            get => _isLoadSaz = FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.enabled", _isLoadSaz);
            set { _isLoadSaz = value; FiddlerApplication.Prefs.SetBoolPref("extensions.EXOFiddlerExtension.enabled", value); }
        }

        private static bool _columnsAllEnabled;
        public static bool ColumnsAllEnabled
        {
            get => _columnsAllEnabled = FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.enabled", _columnsAllEnabled);
            set { _columnsAllEnabled = value; FiddlerApplication.Prefs.SetBoolPref("extensions.EXOFiddlerExtension.enabled", value); }
        }

        private static Int32 _executionCount;
        public static Int32 ExecutionCount
        {
            get => _executionCount = FiddlerApplication.Prefs.GetInt32Pref("extensions.EXOFiddlerExtension.ExecutionCount", 0);
            set { _executionCount = value; FiddlerApplication.Prefs.SetInt32Pref("extensions.EXOFiddlerExtension.ExecutionCount", value); }
        }

    }
}

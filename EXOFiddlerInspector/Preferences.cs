﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EXOFiddlerInspector
{
    /// <summary>
    /// Why preferences.cs? There are also Fiddler application preferences, why does this exist?
    /// Fiddler application preferences are set with for example:
    ///     FiddlerApplication.Prefs.SetBoolPref("extensions.EXOFiddlerInspector.DemoMode", false);
    /// The preferences here are developer created preferences. These preferences are set here, are not changed
    /// at runtime, and are hardcoded in compiled code.
    /// </summary>
    class Preferences
    {
        /////////////////
        /// <summary>
        /// Developer Demo Mode. If enabled as much domain specific information as possible will be replaced with contoso.com.
        /// Note: This is not much right now, just outputs in response comments on the response inspector tab.
        /// </summary>
        Boolean DeveloperDemoMode = false;
        Boolean DeveloperDemoModeBreakScenarios = false;
        /////////////////

        int SlowRunningSessionThreshold = 5000; // milliseconds.

        /// <summary>
        /// Developer list and return.
        /// </summary>
        List<string> Developers = new List<string>(new string[] { "jeknight", "brandev", "bever", "jasonsla" });
        public List<string> GetDeveloperList()
        {
            return Developers;
        }

        /// <summary>
        /// Return DeveloperDemoMode value.
        /// </summary>
        /// <returns></returns>
        public Boolean GetDeveloperMode()
        {
            return DeveloperDemoMode;
        }

        /// <summary>
        /// Return DeveloperDemoModeBreakScenarios value.
        /// </summary>
        /// <returns></returns>
        public Boolean GetDeveloperDemoModeBreakScenarios()
        {
            return DeveloperDemoModeBreakScenarios;
        }

        public int GetSlowRunningSessionThreshold()
        {
            return SlowRunningSessionThreshold;
        }
    }
}

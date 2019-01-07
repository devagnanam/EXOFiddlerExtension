using Fiddler;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EXOFiddlerInspector.Services
{
    /// <summary>
    /// Global application initializer.
    /// </summary>
    public abstract class ActivationService : IAutoTamper
    {
        /// <summary>
        /// This should be consider the main constructor for the app. It's called after the UI has loaded.
        /// </summary>
        public async void OnLoad()
        {   
            MenuUI.Instance.Initialize();
            ColumnsUI.Instance.Initialize();
            SessionProcessor.Instance.Initialize();


            FiddlerApplication.UI.lvSessions.AddBoundColumn("Response Server", 0, 130, "X-ResponseServer");
            FiddlerApplication.UI.lvSessions.AddBoundColumn("Host IP", 0, 110, "X-HostIP");
            FiddlerApplication.UI.lvSessions.AddBoundColumn("Authentication", 0, 140, "X-Authentication");
            FiddlerApplication.UI.lvSessions.AddBoundColumn("Exchange Type", 0, 150, "X-ExchangeType");

            //FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Response Server", 5, -1);
            //FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Host IP", 4, -1);
            //FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Authentication", 3, -1);
            //FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Exchange Type", 2, -1);
            //FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Result", 1, -1);
            //FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("#", 0, -1);


            //FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Elapsed Time", 2, -1);


            // Throw a message box to alert demo mode is running.
            if (Preferences.GetDeveloperMode())
            {
               // MessageBox.Show("Developer / Demo mode is running!");
            }
            else
            {
                await TelemetryService.InitializeAsync();
            }

        }

        public async void OnBeforeUnload()
        {
            await TelemetryService.FlushClientAsync();
        }

        /// <summary>
        /// Called for each HTTP/HTTPS request after it's complete.
        /// </summary>
        /// <param name="_session"></param>
        public void AutoTamperRequestAfter(Session _session) { }

        /// <summary>
        /// Called for each HTTP/HTTPS request before it's complete.
        /// </summary>
        /// <param name="_session"></param>
        public void AutoTamperRequestBefore(Session _session) { }

        /// <summary>
        /// Called for each HTTP/HTTPS response after it's complete.
        /// </summary>
        /// <param name="_session"></param>
        public void AutoTamperResponseAfter(Session _session)
        {
            if (!Preferences.ExtensionEnabled)
            {
                return;
            }

            SessionProcessor.Instance.OnPeekAtResponseHeaders(_session);
            _session.RefreshUI();

            // Call the function to populate the session type column on live trace, if the column is enabled.
            if (Preferences.ExchangeTypeColumnEnabled)
            {
                SessionProcessor.Instance.SetExchangeType(_session);
            }

            // Call the function to populate the session type column on live trace, if the column is enabled.
            if (Preferences.ResponseServerColumnEnabled)
            {
                SessionProcessor.Instance.SetResponseServer(_session);
            }

            // Call the function to populate the Authentication column on live trace, if the column is enabled.
            if (Preferences.AuthColumnEnabled)
            {
                SessionProcessor.Instance.SetAuthentication(_session);
            }
        }

        /// <summary>
        /// Called for each HTTP/HTTPS response before it's complete.
        /// </summary>
        /// <param name="_session"></param>
        public void AutoTamperResponseBefore(Session _session) { }

        /// <summary>
        /// Called for each HTTP/HTTPS error response before it's complete.
        /// </summary>
        /// <param name="_session"></param>
        public void OnBeforeReturningError(Session _session) { }      
    }
}

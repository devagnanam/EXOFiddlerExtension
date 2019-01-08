using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EXOFiddlerInspector.Services;
using Fiddler;

namespace EXOFiddlerInspector
{
    public class ColumnsUI : ActivationService
    {
        private static ColumnsUI _instance;

        public static ColumnsUI Instance => _instance ?? (_instance = new ColumnsUI());

        public bool IsInitialized { get; set; }

        public ColumnsUI()
        {

        }

        public void Initialize()
        {
            this.AddAllEnabledColumns();
            IsInitialized = true;
        }

        public void AddAllEnabledColumns()
        {
            //this.SetElapsedTimeColumn();
            //this.SetResponseServerColumn();
            //this.SetHostIPColumn();
            //this.SetExchangeTypeColumn();
            //this.SetAuthColumn();

            OrderColumns();
        }


        /// <summary>
        /// Set the Response Time Column has been created, return if it has.
        /// </summary>
        public void SetElapsedTimeColumn()
        {
            Boolean LoadSaz = FiddlerApplication.Prefs.GetBoolPref("extensions.EXOFiddlerExtension.LoadSaz", false);

            //if (bElapsedTimeColumnCreated) return;

            if (LoadSaz && Preferences.ElapsedTimeColumnEnabled && Preferences.ExtensionEnabled)
            {
                FiddlerApplication.UI.lvSessions.AddBoundColumn("Elapsed Time", 110, "X-ElapsedTime");                
            }
            else if (Preferences.ExtensionEnabled)
            {
                // live trace, don't load this column.
                // Testing.
                //FiddlerApplication.UI.lvSessions.AddBoundColumn("Elapsed Time", 2, 110, "X-ElapsedTime");
                //bElapsedTimeColumnCreated = true;
            }
        }

        /// <summary>
        ///  Set the Response Server column has been created, return if it has.
        /// </summary>
        public void SetResponseServerColumn()
        {
            if (Preferences.ResponseServerColumnEnabled && !(FiddlerApplication.UI.lvSessions.Columns.Contains((new ColumnHeader() { Text = "X-ResponseServer" }))))
            {
                FiddlerApplication.UI.lvSessions.AddBoundColumn("Response Server", 130, "X-ResponseServer");
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Response Server", 2, -1);
            }            
        }

        /// <summary>
        ///  Set the HostIP column has been created, return if it has.
        /// </summary>
        public void SetHostIPColumn()
        {
            if (Preferences.HostIPColumnEnabled && !(FiddlerApplication.UI.lvSessions.Columns.Contains((new ColumnHeader() { Text = "X-HostIP" }))))
            {
                FiddlerApplication.UI.lvSessions.AddBoundColumn("Host IP", 110, "X-HostIP");
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Host IP", 2, -1);
            }
        }

        /// <summary>
        /// Set the Exchange Type Column has been created, return if it has.
        /// </summary>
        public void SetExchangeTypeColumn()
        {
            if (Preferences.ExchangeTypeColumnEnabled && !(FiddlerApplication.UI.lvSessions.Columns.Contains((new ColumnHeader() { Text = "Exchange Type" }))))
            {
                FiddlerApplication.UI.lvSessions.AddBoundColumn("Exchange Type", 2, 150, "X-ExchangeType");
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Exchange Type", 2, -1);
            }
        }

        public void SetAuthColumn()
        {
            if (Preferences.AuthColumnEnabled && !(FiddlerApplication.UI.lvSessions.Columns.Contains((new ColumnHeader() { Text = "X-Authentication" }))))
            {
                FiddlerApplication.UI.lvSessions.AddBoundColumn("Authentication", 140, "X-Authentication");
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Authentication", 2, -1);
            }
        }

        public void OrderColumns()
        {
            if (Preferences.ExtensionEnabled)
            {
                // Keep session id and result in the standard location on the left.
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("#", 0, -1);
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Result", 1, -1);

                // Set extension added columns here all with a calue of 2 to account for some being enabled.
                if (Preferences.ResponseServerColumnEnabled && !(FiddlerApplication.UI.lvSessions.Columns.Contains((new ColumnHeader() { Text = "X-ResponseServer" }))))
                {
                    FiddlerApplication.UI.lvSessions.AddBoundColumn("Response Server", 130, "X-ResponseServer");
                    FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Response Server", 2, -1);
                }
                if (Preferences.HostIPColumnEnabled && !(FiddlerApplication.UI.lvSessions.Columns.Contains((new ColumnHeader() { Text = "X-HostIP" }))))
                {
                    FiddlerApplication.UI.lvSessions.AddBoundColumn("Host IP", 110, "X-HostIP");
                    FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Host IP", 2, -1);
                }
                if (Preferences.AuthColumnEnabled && !(FiddlerApplication.UI.lvSessions.Columns.Contains((new ColumnHeader() { Text = "X-Authentication" }))))
                {
                    FiddlerApplication.UI.lvSessions.AddBoundColumn("Authentication", 140, "X-Authentication");
                    FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Authentication", 2, -1);
                }
                if (Preferences.ExchangeTypeColumnEnabled && !(FiddlerApplication.UI.lvSessions.Columns.Contains((new ColumnHeader() { Text = "Exchange Type" }))))
                {
                    FiddlerApplication.UI.lvSessions.AddBoundColumn("Exchange Type", 2, 150, "X-ExchangeType");
                    FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Exchange Type", 2, -1);
                }
                if (Preferences.ElapsedTimeColumnEnabled)
                {
                    FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Elapsed Time", 2, -1);
                }
                int iColumnsCount = FiddlerApplication.UI.lvSessions.Columns.Count;
                // Tack the rest on the end using iColumnsCount to avoid out of bounds errors when some columns are disabled.
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Process", iColumnsCount - 9, -1);
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Protocol", iColumnsCount - 8, -1);
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Host", iColumnsCount - 7, -1);
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("URL", iColumnsCount - 6, -1);
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Body", iColumnsCount - 5, -1);
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Caching", iColumnsCount - 4, -1);
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Content-Type", iColumnsCount - 3, -1);
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Comments", iColumnsCount - 2, -1);
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Custom", iColumnsCount - 1, -1);
            }
            // If the extension is disabled, return the UI to the defaults.
            else
            {
                // Move the process column back to its standard position when extension is not enabled.
                FiddlerApplication.UI.lvSessions.SetColumnOrderAndWidth("Process", FiddlerApplication.UI.lvSessions.Columns.Count - 2, -1);
            }

        }

     

        // Populate the ElapsedTime column on live trace, if the column is enabled.
        // Code currently not used / under review.

        // if (boolElapsedTimeColumnEnabled && boolExtensionEnabled) {
        // Realised this.session.oResponse.iTTLB.ToString() + "ms" is not the value I want to display as Response Time.
        // More desirable figure is created from:
        // Math.Round((this.session.Timers.ClientDoneResponse - this.session.Timers.ClientBeginRequest).TotalMilliseconds)
        // 
        // For some reason in AutoTamperResponseAfter this.session.Timers.ClientDoneResponse has a default timestamp of 01/01/0001 12:00
        // Messing up any math. By the time the inspector gets to loading the same math.round statement the correct value is displayed in the 
        // inspector Exchange Online tab.
        //
        // This needs more thought, read through Fiddler book some more on what could be happening and whether this can work or if the Response time
        // column is removed from the extension in favour of the response time on the inspector tab.
        //

        // *** For the moment disabled the Response Time column when live tracing. Only displayed on LoadSAZ. ***

        /*
        // Trying out delaying the process, waiting for the ClientDoneResponse to be correctly populated.
        // Did not work out, Fiddler process hangs / very slow.
        while (this.session.Timers.ClientDoneResponse.Year < 2000)
        {
            if (this.session.Timers.ClientDoneResponse.Year > 2000)
            {
                break;
            }
        }
        //session["X-iTTLB"] = this.session.oResponse.iTTLB.ToString() + "ms"; // Known to give inaccurate results.

        //MessageBox.Show("ClientDoneResponse: " + this.session.Timers.ClientDoneResponse + Environment.NewLine + "ClientBeginRequest: " + this.session.Timers.ClientBeginRequest
        //    + Environment.NewLine + "iTTLB: " + this.session.oResponse.iTTLB);
        // The below is not working in a live trace scenario. Reverting back to the previous configuration above as this works for now.
        session["X-iTTLB"] = Math.Round((this.session.Timers.ClientDoneResponse - this.session.Timers.ClientBeginRequest).TotalMilliseconds) + "ms";
        */
        //}
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Configuration;


namespace ServiceInvioMail
{
    public partial class ServiceInvioMail : ServiceBase
    {
        private int eventId = 1;

        public ServiceInvioMail()
        {
            InitializeComponent();
            //eventLog1 = new System.Diagnostics.EventLog();
            //if (!System.Diagnostics.EventLog.SourceExists("MySource"))
            //{
            //    System.Diagnostics.EventLog.CreateEventSource("MySource", "MyNewLog");
            //}
            //eventLog1.Source = "MySource";
            //eventLog1.Log = "MyNewLog";
        }

        protected override void OnStart(string[] args)
        {
            DateTime now = DateTime.Now;
            if (now.Hour > 6 && now.Hour < 20)
            {
                var timerClock = System.Configuration.ConfigurationManager.AppSettings["timerClock"];

                Timer timer = new Timer();
                timer.Interval = Convert.ToDouble(timerClock); // 3600 seconds = 1 Ora
                timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
                timer.Start();
            }
            //eventLog1.WriteEntry("In OnStart.");
        }


        protected override void OnStop()
        {
            // eventLog1.WriteEntry("In OnStop.");

        }


        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            // TODO: Insert monitoring activities here.
            //eventLog1.WriteEntry("Monitoring the System", EventLogEntryType.Information, eventId++);
            DateTime now = DateTime.Now;
            if (now.Hour > 6 && now.Hour < 20)
            {
                try
                {
                    var myConnect = System.Configuration.ConfigurationManager.AppSettings["DbConnection"];

                    var r = new MailUtilityCon();
                    r.ControlloCarenzeMagazzino(myConnect,"");
                }
                catch (Exception)
                {

                }
            }

        }




    }
}

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
        private static string conNection = "Data Source=10.10.9.4;database=AVdbProd; Integrated Security=false;User ID=UserNetwork;Password=Server2017;";


        public ServiceInvioMail()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                var r = new MailUtilityCon();
                r.ControlloCarenzeMagazzino(conNection, "");
            }
            catch (Exception)
            {

            }

            Timer timer = new Timer();
            timer.Interval = Convert.ToDouble(360000); // 360000 seconds = 1 Ora
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();

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
                    var r = new MailUtilityCon();
                    r.ControlloCarenzeMagazzino(conNection, "");
                }
                catch (Exception)
                {

                }
            }

        }




    }
}

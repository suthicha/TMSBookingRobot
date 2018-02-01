using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using TMSBookingRobot.Controllers;

namespace TMSBookingRobot.WinService
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            timer1.Interval = 3000;
            try
            {
                timer1.Interval = Convert.ToInt32(ConfigurationManager.AppSettings["interval"]);
                timer1.Start();
                Logger.EventLog("Start Service");
            }
            catch { }
        }

        protected override void OnStop()
        {
            timer1.Stop();
            Logger.EventLog("Stop Service");
        }

        private void timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            timer1.Enabled = false;
            try
            {
                var sqlConnectionString = ConfigurationManager.AppSettings["connection"];
                var destFolder = ConfigurationManager.AppSettings["destFolder"];

                var autobot = new BookingAutoBot(sqlConnectionString, destFolder);
                autobot.Live();
            }
            catch (Exception ex)
            {
                Logger.ErrorLog("[Service1.cs][timer1_Elapsed]" + ex.Message);
            }
            timer1.Enabled = true;
        }
    }
}
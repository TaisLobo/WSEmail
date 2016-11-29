using System;
using System.ServiceProcess;
using System.Timers;
using System.Diagnostics;
using System.Configuration;
using EmailComponent;
using System.Net.Mail;
using System.Data;
using System.Data.SqlClient;
namespace EmailService
{
    partial class EmailService : ServiceBase
    {
        private Timer scheduleTimer = null;
        private DateTime lastRun;
        private DateTime lastRunFriday;
        private EventLog eventLogMail;
        private bool flag;

#if DEBUG
        public void StartDebug(string[] args)
        {
            //EnviaEmailSexta();
            OnStart(args);   // chama a rotina principal da classe

        }
#endif

        public EmailService()
        {
            InitializeComponent();
            if (!System.Diagnostics.EventLog.SourceExists("EmailSource"))
            {
                System.Diagnostics.EventLog.CreateEventSource("EmailSource", "EmailLog");
            }
            eventLogMail.Source = "EmailSource";
            eventLogMail.Log = "EmailLog";

            scheduleTimer = new Timer(); //60 minutos
            scheduleTimer.Interval = 1 * 60 * 60 * 1000;
            scheduleTimer.Elapsed += new ElapsedEventHandler(scheduleTimer_Elapsed);

        }

        protected override void OnStart(string[] args)
        {
            flag = true;
            lastRun = DateTime.Now;
            lastRunFriday = DateTime.Now;
            scheduleTimer.Start();
            eventLogMail.WriteEntry("Started");
        }

        protected void scheduleTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            DayOfWeek dia = DateTime.Today.DayOfWeek;
            if (flag == true)
            {

                if (dia == DayOfWeek.Monday)
                {
                    ServiceEmailMethod();
                    
                }

                
                
                lastRun = DateTime.Now;
                flag = false;
            }
            else if (flag == false)
            {
                if (dia == DayOfWeek.Monday && lastRun.Day != DateTime.Now.Day)
                {
                    lastRun = DateTime.Now;
                    ServiceEmailMethod();
                }

               
            }

            

        }

       
        private void ServiceEmailMethod()
        {
            try
            {
                //Sending emails
                eventLogMail.WriteEntry("In Sending Email Method");

                DataSet ds = new DataSet();
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["strconn"].ConnectionString))
                {
                    SqlCommand command = new SqlCommand("RetrievingEmails", conn);
                    command.CommandType = CommandType.StoredProcedure;

                    SqlDataAdapter sda = new SqlDataAdapter(command);
                    sda.Fill(ds);
                }


                EmailComp email = new EmailComp();

                email.subject = "Subject";

                email.messageBody = @"Body"; 


                bool result = email.SendEmail(ds.Tables[0].Rows[0]["email"].ToString(), ds.Tables[0].Rows[0]["Nome"].ToString());
                if (result == true)
                {
                    eventLogMail.WriteEntry("Message Sent SUCCESS to - " + ds.Tables[0].Rows[0]["email"].ToString() + " - " + ds.Tables[0].Rows[0]["email"].ToString());
                }
                else
                {
                    eventLogMail.WriteEntry("Message Sent FAILED to - " + ds.Tables[0].Rows[0]["email"].ToString() + " - " + ds.Tables[0].Rows[0]["email"].ToString());
                }


            }
            catch (Exception erro)
            {
                eventLogMail.WriteEntry("Execution error " + erro.Message);
            }
        }

        protected override void OnStop()
        {
            scheduleTimer.Stop();
            eventLogMail.WriteEntry("Stopped");
        }
        protected override void OnPause()
        {
            scheduleTimer.Stop();
            eventLogMail.WriteEntry("Paused");
        }
        protected override void OnContinue()
        {
            scheduleTimer.Start(); ;
            eventLogMail.WriteEntry("Continuing");
        }
        protected override void OnShutdown()
        {
            scheduleTimer.Stop();
            eventLogMail.WriteEntry("ShutDowned");
        }

        


    }
}

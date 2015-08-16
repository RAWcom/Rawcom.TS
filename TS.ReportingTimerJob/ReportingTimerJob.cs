using System;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Workflow;
using System.Collections;

namespace RAWcom.TS
{
    class ReportingTimerJob : SPJobDefinition
    {
        public static void CreateTimerJob(SPWeb web)
        {
            var timerJob = new ReportingTimerJob(web);

            //timerJob.Schedule = new SPHourlySchedule()
            //{
            //    BeginMinute = 0,
            //    EndMinute = 15
            //};

            //timerJob.Schedule = new SPDailySchedule()
            //{
            //    BeginHour = 20,
            //    BeginMinute = 0,
            //    EndHour = 20,
            //    EndMinute = 15
            //};

            timerJob.Schedule = new SPWeeklySchedule()
            {
                BeginDayOfWeek = DayOfWeek.Friday,
                EndDayOfWeek = DayOfWeek.Friday,
                BeginHour = 21,
                BeginMinute = 0,
                EndHour = 21,
                EndMinute = 0
            };

            timerJob.Update();
        }

        public static void DelteTimerJob(SPWeb web)
        {
            web.Site.WebApplication.JobDefinitions
                .OfType<ReportingTimerJob>()
                .Where(i => string.Equals(i.WebUrl, web.Url, StringComparison.InvariantCultureIgnoreCase))
                .ToList()
                .ForEach(i => i.Delete());
        }

        public ReportingTimerJob()
            : base()
        {

        }

        public ReportingTimerJob(SPWeb web)
            : base(string.Format("ST_Reporting Timer Job ({0})", web.Url), web.Site.WebApplication, null, SPJobLockType.Job)
        {
            Title = Name;
            WebUrl = web.ServerRelativeUrl;
            SiteUrl = web.Site.Url;
            WebID = web.ID;
        }

        public Guid WebID
        {
            get { return (Guid)this.Properties["WebID"]; }
            set { this.Properties["WebID"] = value; }
        }

        public string WebUrl
        {
            get { return (string)this.Properties["WebUrl"]; }
            set { this.Properties["WebUrl"] = value; }
        }

        public string SiteUrl
        {
            get { return (string)this.Properties["SiteUrl"]; }
            set { this.Properties["SiteUrl"] = value; }
        }

        public override void Execute(Guid targetInstanceId)
        {

            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                try
                {
                    using (var site = new SPSite(SiteUrl))
                    {
                        using (var web = site.OpenWeb(WebUrl))
                        {
                            //string webUrl = web.Url;
                            DateTime hitDate = DateTime.Today;

                            ArrayList activeCustomers = RAWcom.TS.Report.GetActiveCustomers(hitDate, web);

                            if (activeCustomers != null && activeCustomers.Count > 0)
                            {
                                foreach (int custId in activeCustomers)
                                {
                                    RAWcom.TS.Report.CreateReportForCustomer(hitDate, custId, web);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {

                    RAWcom.TS.ElasticTestMail.ReportErrorViaEmail(ex, WebUrl);

                }
            });
        }

    }

}


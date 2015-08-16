using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Linq;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using System.Collections;

namespace ConsoleApplication1
{
    class Program
    {
        /// <summary>
        /// CreateReport
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            string siteUrl = @"http://spf2010/sites/rawcom";
            string webUrl = @"/sites/rawcom/TS";
            //DateTime hitDate = DateTime.Today;
            DateTime hitDate = new DateTime(2015, 7, 31);

            using (var site = new SPSite(siteUrl))
            {
                using (var web = site.OpenWeb(webUrl))
                {
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
    }
}

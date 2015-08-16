using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using System.Collections;

namespace RaportEventReceiver.EventReceiver1
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class EventReceiver1 : SPItemEventReceiver
    {
       /// <summary>
       /// An item was added.
       /// </summary>
       public override void ItemAdded(SPItemEventProperties properties)
       {
           Execute(properties);
       }

       private void Execute(SPItemEventProperties properties)
       {
           this.EventFiringEnabled = false;

           try
           {

                   //określ rodzaj raportu
                   string ct = properties.ListItem["ContentType"].ToString();
                   switch (ct)
                   {
                       case "Zestawienie prac serwisowych":
                           properties.ListItem["colStatus"] = "W trakcie obsługi";
                           properties.ListItem.Update();

                           DateTime targetDate = (DateTime)properties.ListItem["colTargetDate"];

                           CreateReport(properties, targetDate);

                           //Create_ZestawieniePracSerwisowych(properties, targetDate);

                           properties.ListItem["colStatus"] = "Zakończony";
                           properties.ListItem.Update();
                           break;
                       default:
                           properties.ListItem["colStatus"] = "Anulowany";
                           properties.ListItem.Update();
                           break;
                   }

           }
           catch (Exception ex)
           {
               properties.ListItem["colStatus"] = "Anulowany";
               properties.ListItem.Update();
               string result = RAWcom.TS.ElasticTestMail.ReportErrorViaEmail(ex,properties.Web.Url);
           }

           this.EventFiringEnabled = true;
       }

       private void CreateReport(SPItemEventProperties properties, DateTime targetDate)
       {

           string SiteUrl = properties.Web.Site.Url;
           string WebUrl = properties.Web.ServerRelativeUrl;

           using (var site = new SPSite(SiteUrl))
           {
               using (var web = site.OpenWeb(WebUrl))
               {

                   ArrayList activeCustomers = RAWcom.TS.Report.GetActiveCustomers(targetDate, web);

                   if (activeCustomers != null && activeCustomers.Count > 0)
                   {
                       foreach (int custId in activeCustomers)
                       {
                           RAWcom.TS.Report.CreateReportForCustomer(targetDate, custId, web);
                       }
                   }
               }
           }
       }

       private void Create_ZestawieniePracSerwisowych(SPItemEventProperties properties, DateTime targetDate)
       {
           string SiteUrl = properties.Web.Site.Url;
           string WebUrl = properties.Web.Url;

           try
           {
               using (var site = new SPSite(SiteUrl))
               {
                   //using (var web = site.OpenWeb(WebUrl))
                   using (var web = site.AllWebs[WebUrl])
                   {

                       ArrayList activeCustomers = RAWcom.TS.Report.GetActiveCustomers(targetDate, web);

                       if (activeCustomers != null && activeCustomers.Count > 0)
                       {
                           foreach (int custId in activeCustomers)
                           {
                               RAWcom.TS.Report.CreateReportForCustomer(targetDate, custId, web);
                           }
                       }
                   }
               }
           }
           catch (Exception ex)
           {

               RAWcom.TS.ElasticTestMail.ReportErrorViaEmail(ex, WebUrl);

           }
       }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using Microsoft.SharePoint;

namespace RAWcom.TS
{
    public class Report
    {
        public static ArrayList GetActiveCustomers(DateTime hitDate, SPWeb web)
        {
            ArrayList results = new ArrayList();

            DateTime baseDate = new DateTime(hitDate.Year, hitDate.Month, 1);


                    var targetList = web.Lists.TryGetList("Karty pracy");

                    if (targetList != null)
                    {
                        targetList.Items.Cast<SPListItem>()
                            .Where(i => (DateTime)i["colData"] >= baseDate)
                            .Where(i => (DateTime)i["colData"] <= hitDate)
                            .Where(i => i["colCzyRozliczony"] == null || (bool)i["colCzyRozliczony"] != true)
                            .GroupBy(i => i["selKlient"])
                            .ToList()
                            .ForEach(item =>
                            {
                                string groupItemKey = item.Key.ToString();
                                int customerId = new SPFieldLookupValue(groupItemKey).LookupId;

                                results.Add(customerId);
                            });
                    }


                return results;

        }

        public static void CreateReportForCustomer(DateTime hitDate, int customerId, SPWeb web)
        {

            string TBodyFormat = @"<P style=""MARGIN-BOTTOM: 1em; FONT-SIZE: 12px; FONT-FAMILY: Arial, Helvetica, sans-serif; COLOR: #a7a7a7; MARGIN-TOP: 0px; BACKGROUND-COLOR: transparent"" align=left>{0}</P>";
            string THeaderFormat = @"<P style=""MARGIN-BOTTOM: 1em; FONT-SIZE: 14px; FONT-FAMILY: Arial, Helvetica, sans-serif; COLOR: #a8a7a7; MARGIN-TOP: 0px; BACKGROUND-COLOR: transparent"" align=left><STRONG>{0}</STRONG></P>";

            DateTime baseDate = new DateTime(hitDate.Year, hitDate.Month, 1);

                    var targetList = web.Lists.TryGetList("Karty pracy");

                    if (targetList != null)
                    {
                        Customer customer = new Customer(customerId, web);

                        StringBuilder sbtr = new StringBuilder();

                        TimeSpan totalMinutes = TimeSpan.FromMinutes(0);

                        targetList.Items.Cast<SPListItem>()
                            .Where(i => (DateTime)i["colData"] >= baseDate)
                            .Where(i => (DateTime)i["colData"] <= hitDate)
                            .Where(i => i["colCzyRozliczony"] == null || (bool)i["colCzyRozliczony"] != true)
                            .Where(i => new SPFieldLookupValue(i["selKlient"] as string).LookupId == customerId)
                            //.OrderByDescending(i => i.ID)
                            //.ThenByDescending(i=>i["colData"])
                            .OrderBy(i => i["colData"])
                            .ThenBy(i => i.ID)
                            .ToList()
                            .ForEach(item =>
                            {
                                //update totalMinutes
                                if (item["colCzasMin"] != null)
                                {
                                    int minutes = 0;
                                    int.TryParse(item["colCzasMin"].ToString(), out minutes);

                                    if (minutes > 0)
                                    {
                                        totalMinutes = totalMinutes.Add(TimeSpan.FromMinutes(minutes));
                                    }
                                }

                                //create report line
                                StringBuilder tr = new StringBuilder(@"<tr>
		<td>___ID___</td>
		<td>___Data___</td>
		<td>___Temat___</td>
		<td>___Czas___</td>
		<td>___Status___</td>
	</tr>");
                                tr.Replace("___ID___", String.Format(TBodyFormat, item.ID.ToString()));
                                tr.Replace("___Data___", String.Format(TBodyFormat, ((DateTime)item["colData"]).ToString("MM-dd")));
                                tr.Replace("___Temat___", String.Format(TBodyFormat, (item["Title"]).ToString()));
                                tr.Replace("___Czas___", String.Format(TBodyFormat, item["colCzasMin"] != null ? item["colCzasMin"].ToString() + " min" : ""));
                                //tr.Replace("___Koszt___", String.Format(TBodyFormat,item["colCzasMin"] != null ? item["colCzasMin"].ToString() : ""));
                                tr.Replace("___Rozliczony?___", String.Format(TBodyFormat, item["colCzyRozliczony"] == null || (bool)item["colCzyRozliczony"] != true ? "Nie" : "Tak"));
                                tr.Replace("___Status___", String.Format(TBodyFormat, item["Status"] != null ? item["Status"].ToString() : ""));

                                sbtr.Append(tr);

                                if (item["colOpis"] != null)
                                {
                                    tr = new StringBuilder(@"<tr><td>&nbsp;</td><td>&nbsp;</td><td colspan=""2"">___Opis___</td><td>&nbsp;</td></tr>");
                                    tr.Replace("___Opis___", String.Format(TBodyFormat, item["colOpis"] != null ? item["colOpis"].ToString() : ""));
                                    sbtr.Append(tr);
                                }

                            });

                        StringBuilder sb = new StringBuilder(@"<table>
	<tr>
		<td>___ID___</td>
		<td>___Data___</td>
		<td>___Temat___</td>
		<td>___Czas___</td>
		<td>___Status___</td>
	</tr>
___Rows___
</table>");

                        sb.Replace(@"<table>", @"<table align=""left"" border=""0"" cellpadding=""2"" cellspacing=""0"" style=""width: 100.0%; border-collapse: collapse;  margin-left: -1.8pt; margin-right: -1.8pt; font-size: 10.0pt; font-family: Arial, Helvetica, sans-serif;"" width=""100%"">");
                        sb.Replace(@"<td>", @"<td>");
                        sb.Replace("___ID___", String.Format(THeaderFormat, "ID#"));
                        sb.Replace("___Data___", String.Format(THeaderFormat, "Data"));
                        sb.Replace("___Temat___", String.Format(THeaderFormat, "Temat"));
                        //sb.Replace("___Opis___", String.Format(THeaderFormat, "Opis"));
                        sb.Replace("___Czas___", String.Format(THeaderFormat, "Czas"));
                        sb.Replace("___Koszt___", String.Format(THeaderFormat, "Koszt"));
                        sb.Replace("___Status___", String.Format(THeaderFormat, "Status"));

                        sb.Replace("___Rows___", sbtr.ToString());

                        string subject = String.Format("::Zestawienie prac serwisowych za okres {0} [{1}]",
                            hitDate.ToString("yyyy-MM"),
                            customer.Name.ToString());
                        string bodyHTML = sb.ToString();

                        double totalHoursToReimburse = totalMinutes.TotalHours;

                        UpdateReport(baseDate, hitDate, customer, subject, bodyHTML, web, totalHoursToReimburse);

                    }
        }

        public static void UpdateReport(DateTime baseDate, DateTime hitDate, Customer customer, string subject, string bodyHTML, SPWeb web, Double totalHoursToReimburse)
        {

            var targetList = web.Lists.TryGetList("Raporty");
            if (targetList != null)
            {
                SPListItem report = (targetList.Items.Cast<SPListItem>()
                                        .Where(i => new SPFieldLookupValue(i["selKlient"] as string).LookupId == customer.ID)
                                        .Where(i => ((DateTime)i["colReportingDate"]).ToShortDateString() == hitDate.ToShortDateString())
                                        .Where(i => (i["colIsSent"] == null || (bool)i["colIsSent"] != true))
                                        .FirstOrDefault());



                if (report == null)
                {
                    report = targetList.AddItem();
                    report["selKlient"] = customer.ID;
                    report["colReportingDate"] = hitDate;
                    report["colBaseDate"] = baseDate;
                }

                report["colSubject"] = subject;
                report["colTo"] = customer.Email;
                report["colCc"] = customer.Cc;
                report["colBodyHTML"] = bodyHTML;
                report["colTotalHoursToReimburse"] = Math.Round(totalHoursToReimburse, 2);
                report["colScheduledDeliveryDate"] = DateTime.Now;

                report["Title"] = DateTime.Now.ToString();
                report["colIsSent"] = false;


                report.Update();

                Helpers.StartWorkflow(report, "Wysyłka Raportu");

            }
        }

    }
}

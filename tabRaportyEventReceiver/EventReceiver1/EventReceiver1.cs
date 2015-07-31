using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using System.Linq;
using Microsoft.SharePoint.Linq;
using System.Diagnostics;

namespace tabRaportyEventReceiver.EventReceiver1
{
    /// <summary>
    /// tabKontrakty
    /// </summary>
    public class EventReceiver1 : SPItemEventReceiver
    {
        public override void ItemAdded(SPItemEventProperties properties)
        {
            Execute(properties);
        }

        public override void ItemUpdated(SPItemEventProperties properties)
        {
            Execute(properties);
        }

        private void Execute(SPItemEventProperties properties)
        {
            this.EventFiringEnabled = false;
            
            //SPSecurity.RunWithElevatedPrivileges(delegate()
            //{

                Execute_Main(properties);

            //});
            
            this.EventFiringEnabled = true;
	
        }

        private static void Execute_Main(SPItemEventProperties properties)
        {
            using (SPSite site = properties.Web.Site)
            {
                using (TSDataContext ctx = new TSDataContext(properties.Web.Site.Url))
                {






                    ctx.ObjectTrackingEnabled = true;

                    var results = from r in ctx.KartyPracy.OfType<KartyPracyKartaPracy>()
                                  where r.DataRejestracji == DateTime.Today
                                  && !String.IsNullOrEmpty(r.Klient.Email)
                                  && r.Rozliczony != true
                                  orderby r.Klient.NazwaKlienta
                                  orderby r.DataRejestracji
                                  select r;
                    //select new
                    //{
                    //    ID = r.Id,
                    //    Klient = r.Klient.NazwaKlienta,
                    //    Data = r.DataRejestracji,
                    //    Czas = r.Czas,
                    //    Email = r.Klient.Email,
                    //    Rozliczony = r.Rozliczony,
                    //    Uwagi = r.Uwagi
                    //};


                    //string memo = DateTime.Now.ToString() + " ";

                    //foreach (var item in results)
                    //{
                    //    //Debug.WriteLine(String.Format("ID={0}, Klient={1}, Data={2}, Czas={3}, Email={4}",
                    //    //    item.ID,
                    //    //    item.Klient,
                    //    //    ((DateTime)item.Data).ToShortDateString(),
                    //    //    item.Czas,
                    //    //    item.Email));

                    //    //KartyPracyKartaPracy kp = ctx.KartyPracy.Where(r => r.Id == item.ID).FirstOrDefault();

                    //    //kp.OstatniaAktualizacja = memo + item.Email;
                    //    //if (kp.Rozliczony != true)
                    //    //{
                    //    //    kp.Rozliczony = false;
                    //    //}

                    //}

                    //ctx.SubmitChanges();
                }
            }
        }
    }
}
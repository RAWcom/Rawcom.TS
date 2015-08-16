using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace RAWcom.TS
{
    public class Customer
    {
        private int customerId;
        private string siteUrl;
        private Guid webId;

        public Customer(int customerId, SPWeb web)
        {
            // TODO: Complete member initialization
            this.customerId = customerId;

            var targetList = web.Lists.TryGetList("Klienci");
            if (targetList != null)
            {
                SPListItem customer = (targetList.Items.Cast<SPListItem>()
                                        .Where(i => i.ID == customerId)
                                        .FirstOrDefault());
                if (customer != null)
                {
                    this.ID = customerId;
                    this.Name = customer["colNazwaKlienta"].ToString();
                    this.Email = customer["colEmail"].ToString();
                    if (customer["colCC"] != null)
                    {
                        this.Cc = customer["colCC"].ToString();
                    }
                }
            }
        }
        public int ID { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string Cc { get; set; }
    }
}

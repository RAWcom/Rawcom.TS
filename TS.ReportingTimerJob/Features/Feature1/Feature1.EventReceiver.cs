using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;

namespace RAWcom.TS.Features.Feature1
{
    [Guid("0bfb2887-1940-49f2-be75-98735880a2b7")]
    public class Feature1EventReceiver : SPFeatureReceiver
    {
        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            var web = properties.Feature.Parent as SPWeb;
            RAWcom.TS.ReportingTimerJob.CreateTimerJob(web);
        }

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            var web = properties.Feature.Parent as SPWeb;
            RAWcom.TS.ReportingTimerJob.DelteTimerJob(web);
        }
    }
}

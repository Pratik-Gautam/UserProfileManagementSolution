using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;

namespace SPOFunctionApp
{
    public static class UpdateUserProfile
    {
        [FunctionName("UpdateUserProfile")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("Triggering Azure Function to process the JSON");

            // parse query parameter
            string fileName = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "fileName", true) == 0)
                .Value;

            string tenantAdminUrl = "https://m365x131504-admin.sharepoint.com";
            // User name and pwd to login to the tenant
            string userName = "admin@m365x131504.onmicrosoft.com";
            string pwd = "SPF99BO4kF";
            
            string fileUrl = "https://m365x131504.sharepoint.com/sites/PowerApps/Shared%20Documents/" + fileName;
            ClientResult<Guid> workItemId;
            string status = string.Empty;

            // Get access to source tenant with tenant permissions
            using (var ctx = new ClientContext(tenantAdminUrl))
            {
                //Provide count and pwd for connecting to the source
                var passWord = new SecureString();
                foreach (char c in pwd.ToCharArray()) passWord.AppendChar(c);
                ctx.Credentials = new SharePointOnlineCredentials(userName, passWord);

                // Only to check connection and permission, could be removed
                ctx.Load(ctx.Web);
                ctx.ExecuteQuery();
                string title = ctx.Web.Title;

                // Let's get started on the actual code!!!
                Office365Tenant tenant = new Office365Tenant(ctx);
                ctx.Load(tenant);
                ctx.ExecuteQuery();

                /// /// /// /// /// /// /// /// ///
                /// DO import based on file whcih is already uploaded to tenant
                /// /// /// /// /// /// /// /// ///

                // Type of user identifier ["PrincipleName", "EmailAddress", "CloudId"] 
                // in the User Profile Service.
                // In this case we use email as the identifier at the UPA storage
                ImportProfilePropertiesUserIdType userIdType = ImportProfilePropertiesUserIdType.Email;

                // Name of user identifier property in the JSON
                var userLookupKey = "IdName";

                var propertyMap = new System.Collections.Generic.Dictionary<string, string>();
                // First one is the file, second is the target at User Profile Service
                // Notice that we have here 2 custom properties in UPA called 'City' and 'Office'
                propertyMap.Add("MyCustomProperty", "MyCustomProperty");
                propertyMap.Add("MyCustomProperty1", "MyCustomProperty1");

                // Returns a GUID, which can be used to see the status of the execution and end results
                workItemId = tenant.QueueImportProfileProperties(
                                        userIdType, userLookupKey, propertyMap, fileUrl
                                        );

                ctx.ExecuteQuery();

                var job = tenant.GetImportProfilePropertyJob(workItemId.Value);
                ctx.Load(job);
                ctx.ExecuteQuery();

                status=string.Format("ID: {0} - Request status: {1} - Error status: {2}",
                                  job.JobId, job.State.ToString(), job.Error.ToString());
            }


            return workItemId == null
                ? req.CreateResponse(HttpStatusCode.BadRequest, "Some Error Occurred")
                : req.CreateResponse(HttpStatusCode.OK, status);
        }
    }
}

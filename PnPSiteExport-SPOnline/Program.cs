using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Security;

namespace PnPSiteExport_SPOnline
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Site Url (e.g. https://tenant.sharepoint.com/sites/test): ");
            string siteUrl = Console.ReadLine();

            Console.WriteLine("User Name (e.g. user@tenant.onmicrosoft.com): ");
            string username = Console.ReadLine();

            SecureString password = SecurePasswordFetcher.GetPassword();

            Console.WriteLine("File Name (e.g. export.xml): ");
            string filename = Console.ReadLine();

            Console.WriteLine(@"Please ensure a folder exists at C:\PnPSiteExport");
            Console.WriteLine("Press any key to start generating the export.");
            Console.ReadLine();


            var credentials = new SharePointOnlineCredentials(username, password);

            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                clientContext.Credentials = credentials;

                Web web = clientContext.Web;

                clientContext.Load(web);
                clientContext.ExecuteQueryRetry();
                var provisioningInfo = new ProvisioningTemplateCreationInformation(web);
                provisioningInfo.FileConnector = new FileSystemConnector(@"C:\PnPSiteExport", "");
                //provisioningInfo.PersistComposedLookFiles = true;
                provisioningInfo.IncludeAllTermGroups = true;
                provisioningInfo.IncludeSearchConfiguration = true;
                provisioningInfo.IncludeSiteGroups = true;
                provisioningInfo.IncludeTermGroupsSecurity = true;
                var template = web.GetProvisioningTemplate(provisioningInfo);

                XMLFileSystemTemplateProvider provider = new XMLFileSystemTemplateProvider(@"C:\PnPSiteExport", "");
                provider.SaveAs(template, filename);

            }
        }

        public static class SecurePasswordFetcher
        {
            public static SecureString GetPassword()
            {
                Console.Write("Password: ");
                SecureString pwd = new SecureString();
                while (true)
                {
                    ConsoleKeyInfo i = Console.ReadKey(true);
                    if (i.Key == ConsoleKey.Enter)
                    {
                        Console.WriteLine("\n");
                        break;
                    }
                    else if (i.Key == ConsoleKey.Backspace)
                    {
                        if (pwd.Length > 0)
                        {
                            pwd.RemoveAt(pwd.Length - 1);
                            Console.Write("\b \b");
                        }
                    }
                    else
                    {
                        pwd.AppendChar(i.KeyChar);
                        Console.Write("*");
                    }
                }
                return pwd;
            }
        }
    }
}

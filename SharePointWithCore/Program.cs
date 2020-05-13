using Microsoft.SharePoint.Client;
using System;
using System.Net;
using System.Security;

namespace SharePointWithCore
{
    class Program
    {
        static string login = "veaceslav.chirilov@endava.com"; //give your username here  
        static string password = "Spartak%1005"; //give your password  

        //static string siteUrl = "https://endava.sharepoint.com/_layouts/15";
        //static string listName = "NewJoinersList";
        //static string viewName = "All Items";

        static string siteUrl = "https://endava.sharepoint.com/sites/ProjectDBDev/_layouts/15";
        static string listName = "Projects Oracle";
        static string viewName = "All Items";


        

        static void Main(string[] args)
        {
            //FetchListRecords();
            Insert();
            Console.ReadKey();
        }
                
        public static void FetchListRecords()
        {
            ClientContext clientContext = new ClientContext(siteUrl);

            var myList = clientContext.Web.Lists.GetByTitle(listName);

            clientContext.Load(myList);

            var onlineCredentials = new SharePointOnlineCredentials(login, password);
            clientContext.Credentials = onlineCredentials;

            clientContext.ExecuteQueryAsync().Wait();
            View view = myList.Views.GetByTitle(viewName);

            clientContext.Load(view);
            clientContext.ExecuteQueryAsync().Wait();

            // Create a new CAML Query object and store the query from the custom view
            //
            CamlQuery query = new CamlQuery();
            query.ViewXml = view.ViewQuery;

            // Based on the query load items
            //
            ListItemCollection items = myList.GetItems(query);
            clientContext.Load(items);
            clientContext.ExecuteQueryAsync().Wait();
            //Console.WriteLine(items.Count);

            foreach (var item in items)
            {
                Console.WriteLine($"Project Name \t {item["Title"]}");
                Console.WriteLine($"Delivery Lead \t {item["Delivery_x0020_Lead"]}");
            }
        }

        static void Insert()
        {
            try
            {
                using (var clientContext = new ClientContext(siteUrl))
                {
                    var myList = clientContext.Web.Lists.GetByTitle(listName);
                    ListItemCreationInformation itemInfo = new ListItemCreationInformation();

                    ListItem myItem = myList.AddItem(itemInfo);
                    myItem["Title"] = "Auto-Inserted Project";

                    myItem.Update();

                    var onlineCredentials = new SharePointOnlineCredentials(login, password);
                    clientContext.Credentials = onlineCredentials;
                    clientContext.ExecuteQueryAsync().Wait();
                    Console.WriteLine("New Item inserted Successfully");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}

using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace SharePointWithCore
{
    class Program
    {
        static string login = "veaceslav.chirilov@endava.com"; 
        static string password = "Nautilus_1001"; 

        //static string siteUrl = "https://endava.sharepoint.com/_layouts/15";
        //static string listName = "NewJoinersList";
        //static string viewName = "All Items";

        static string siteUrl = "https://endava.sharepoint.com/sites/ProjectDBDev/_layouts/15";
        static string listName = "Projects Oracle";
        static string viewName = "All Items";
        

        static void Main(string[] args)
        {
            //Delete();
            //Console.ReadKey();
            //return;

            //var projects = FetchAllRecords();

            List<ProjectCodeDataItem> items = new List<ProjectCodeDataItem>
            {
                new ProjectCodeDataItem { Project_Name = "GPS", Delivery_Lead = "Chris Martin", Client_Location = "London, UK", Start_Date = DateTime.Now, End_Date = DateTime.Now, Project_Code="GGY223", Project_Description = "Payment", Client_Name="Global Payment System"},
                new ProjectCodeDataItem { Project_Name = "Telecom", Delivery_Lead = "Adam Ferguson", Client_Location = "New York, USA", Start_Date = DateTime.Now, End_Date = DateTime.Now, Project_Code="SFD002", Project_Description = "Telecomunications", Client_Name="Royal Telecom"},
                new ProjectCodeDataItem { Project_Name = "Mastercard", Delivery_Lead = "Paul Brown", Client_Location = "Calofornia, USA", Start_Date = DateTime.Now, End_Date = DateTime.Now, Project_Code="KJH844", Project_Description = "Payment", Client_Name="Mastercad"},
                new ProjectCodeDataItem { Project_Name = "Vocalink", Delivery_Lead = "Eoin Woods", Client_Location = "Paris, France", Start_Date = DateTime.Now, End_Date = DateTime.Now, Project_Code="POY515", Project_Description = "Finance", Client_Name="Vocalink"},
                new ProjectCodeDataItem { Project_Name = "Concardis", Delivery_Lead = "Akin Olushoga", Client_Location = "London, UK", Start_Date = DateTime.Now, End_Date = DateTime.Now, Project_Code="QWA644", Project_Description = "Trading", Client_Name="Concardis"}
            };

            items.Clear();

            for (int i = 1; i <= 300; i++)
            {
                items.Add(new ProjectCodeDataItem { Project_Name = $"Project {i}", Delivery_Lead = $"Delivery Lead {i}", Client_Location = $"Client Location {i}", Start_Date = DateTime.Now, End_Date = DateTime.Now, Project_Code = $"Code-{i}", Project_Description = $"Description {i}", Client_Name = $"Client {i}" });
            }

            Insert(items);
            //Update(items);
            Console.ReadKey();
        }
                
        public static List<ProjectCodeDataItem> FetchAllRecords()
        {
            List<ProjectCodeDataItem> result = new List<ProjectCodeDataItem>();

            using (var clientContext = new ClientContext(siteUrl))
            {
                var myList = clientContext.Web.Lists.GetByTitle(listName);

                clientContext.Load(myList);                
                clientContext.Credentials = new SharePointOnlineCredentials(login, password);

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
                    result.Add(new ProjectCodeDataItem
                    {
                        Project_Name = item["Project_Name"].ToString(),
                        Delivery_Lead = item["Delivery_Lead"].ToString()                        
                    });
                }

                return result;
            }
        }

        static void Insert(List<ProjectCodeDataItem> eis_projects, int batchSize = 100)
        {
            try
            {
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();

                using (var clientContext = new ClientContext(siteUrl))
                {
                    var spList = clientContext.Web.Lists.GetByTitle(listName);

                    clientContext.Credentials = new SharePointOnlineCredentials(login, password);

                    Console.WriteLine($"Total items to be inserted {eis_projects.Count()}");

                    foreach (var batch in eis_projects.Batch(batchSize))
                    {
                        Console.WriteLine($"{batch.Count()} items are being inserted");

                        foreach (var eis_project in batch)
                        {
                            ListItem sp_project = spList.AddItem(new ListItemCreationInformation());

                            sp_project["Project_Name"] = eis_project.Project_Name;
                            sp_project["Delivery_Lead"] = eis_project.Delivery_Lead;
                            sp_project["Start_Date"] = eis_project.Start_Date;
                            sp_project["End_Date"] = eis_project.End_Date;
                            sp_project["Client_Location"] = eis_project.Client_Location;
                            sp_project["Project_Code"] = eis_project.Project_Code;
                            sp_project["Project_Description"] = eis_project.Project_Description;
                            sp_project["Client_Name"] = eis_project.Client_Name;
                            sp_project["Business_Unit_AGU"] = eis_project.Business_Unit_AGU;
                            sp_project["Oracle_Comments"] = eis_project.Oracle_Comments;
                            sp_project["Database_Platform"] = eis_project.Database_Platform;
                            sp_project["Programming_Language"] = eis_project.Programming_Language;
                            sp_project["Infrastructure_Platform"] = eis_project.Infrastructure_Platform;
                            sp_project["Modified"] = eis_project.Modified;
                            sp_project["Created"] = eis_project.Created;

                            sp_project.Update();
                        }
                        
                        clientContext.ExecuteQueryAsync().Wait();
                    }                    
                }

                stopWatch.Stop();                
                TimeSpan ts = stopWatch.Elapsed;
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                Console.WriteLine($"Inserted Successfully in {elapsedTime}");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
                       
              
        static void Update(List<ProjectCodeDataItem> ets_projects)
        {
            try
            {
                using (var clientContext = new ClientContext(siteUrl))
                {                    
                    var sp_list = clientContext.Web.Lists.GetByTitle(listName);

                    clientContext.Credentials = new SharePointOnlineCredentials(login, password);

                    foreach (var eis_project in ets_projects)
                    {
                        CamlQuery query = new CamlQuery();
                        query.ViewXml = @"<View><Query><Where><BeginsWith><FieldRef Name='Project_Name' /><Value Type='Text'>" + eis_project.Project_Name + @"</Value></BeginsWith></Where></Query></View>";
                        ListItemCollection sp_projects = sp_list.GetItems(query);
                        clientContext.Load(sp_projects);
                        clientContext.ExecuteQueryAsync().Wait();


                        if (sp_projects?.Count()>0 && eis_project.Project_Name.Equals((string)sp_projects[0]["Project_Name"], StringComparison.OrdinalIgnoreCase))
                        {
                            sp_projects[0]["Project_Name"] = eis_project.Project_Name;
                            sp_projects[0]["Delivery_Lead"] = eis_project.Delivery_Lead;
                            sp_projects[0]["Start_Date"] = eis_project.Start_Date;
                            sp_projects[0]["End_Date"] = eis_project.End_Date;
                            sp_projects[0]["Client_Location"] = eis_project.Client_Location;
                            sp_projects[0]["Project_Code"] = eis_project.Project_Code;
                            sp_projects[0]["Project_Description"] = eis_project.Project_Description;
                            sp_projects[0]["Client_Name"] = eis_project.Client_Name;
                            sp_projects[0]["Business_Unit_AGU"] = eis_project.Business_Unit_AGU;
                            sp_projects[0]["Oracle_Comments"] = eis_project.Oracle_Comments;
                            sp_projects[0]["Database_Platform"] = eis_project.Database_Platform;
                            sp_projects[0]["Programming_Language"] = eis_project.Programming_Language;
                            sp_projects[0]["Infrastructure_Platform"] = eis_project.Infrastructure_Platform;
                            sp_projects[0]["Modified"] = eis_project.Modified;
                            sp_projects[0]["Created"] = eis_project.Created;

                            sp_projects[0].Update();
                        }
                        else
                        {
                            ListItem sp_project = sp_list.AddItem(new ListItemCreationInformation());
                            sp_project["Project_Name"] = eis_project.Project_Name;
                            sp_project["Delivery_Lead"] = eis_project.Delivery_Lead;
                            sp_project["Start_Date"] = eis_project.Start_Date;
                            sp_project["End_Date"] = eis_project.End_Date;
                            sp_project["Client_Location"] = eis_project.Client_Location;
                            sp_project["Project_Code"] = eis_project.Project_Code;
                            sp_project["Project_Description"] = eis_project.Project_Description;
                            sp_project["Client_Name"] = eis_project.Client_Name;
                            sp_project["Business_Unit_AGU"] = eis_project.Business_Unit_AGU;
                            sp_project["Oracle_Comments"] = eis_project.Oracle_Comments;
                            sp_project["Database_Platform"] = eis_project.Database_Platform;
                            sp_project["Programming_Language"] = eis_project.Programming_Language;
                            sp_project["Infrastructure_Platform"] = eis_project.Infrastructure_Platform;
                            sp_project["Modified"] = eis_project.Modified;
                            sp_project["Created"] = eis_project.Created;

                            sp_project.Update();
                        }                        
                    }

                    clientContext.ExecuteQueryAsync().Wait();
                }

                Console.WriteLine("Updated Successfully");

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        static void Delete(int batchSize = 100)
        {
            try
            {
                using (var clientContext = new ClientContext(siteUrl))
                {
                    var spList = clientContext.Web.Lists.GetByTitle(listName);

                    clientContext.Load(spList);
                    clientContext.Credentials = new SharePointOnlineCredentials(login, password);
                    //clientContext.ExecuteQueryAsync().Wait();

                    View view = spList.Views.GetByTitle(viewName);
                    clientContext.Load(view);
                    clientContext.ExecuteQueryAsync().Wait();

                    var count = spList.ItemCount;
                    var batches = count.ToArray().Batch(batchSize);
                    Console.WriteLine($"Total records to be deleted {count}");

                    foreach(var batch in batches)
                    {
                        Console.WriteLine($"Items {batch.Min()}..{batch.Max()} are being deleted");
                        int innerCount = batch.Count();
                        // Create a new CAML Query object and store the query from the custom view
                        CamlQuery query = new CamlQuery();
                        query.ViewXml = view.ViewQuery;

                        ListItemCollection items = spList.GetItems(query);
                        clientContext.Load(items, y => y.Take(innerCount));
                        clientContext.ExecuteQueryAsync().Wait();

                        items.ToList().ForEach(item => item.DeleteObject());
                        clientContext.ExecuteQueryAsync().Wait();
                    }                    
                }

                Console.WriteLine("All items have been deleted succesfully.");
            }            
            catch(Exception exc)
            {
                Console.WriteLine($"Delete Operation Failed: {exc.Message}");
            }
        }
    }
}

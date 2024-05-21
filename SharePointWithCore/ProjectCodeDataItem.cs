using System;
using System.Collections.Generic;
using System.Text;

namespace SharePointWithCore
{
    class ProjectCodeDataItem
    {
        public ProjectCodeDataItem()
        {
            Project_Name = string.Empty;
            Delivery_Lead = string.Empty;
            Start_Date = DateTime.Now;
            End_Date = DateTime.Now;
            Client_Location = string.Empty;
            Project_Description = string.Empty;
            Client_Name = string.Empty;
            Business_Unit_AGU = string.Empty;
            Oracle_Comments = string.Empty;
            Database_Platform = string.Empty;
            Programming_Language = string.Empty;
            Infrastructure_Platform = string.Empty;
            Modified = DateTime.Now;
            Created = DateTime.Now;

        }
        public string Project_Name { get; set; }
        public string Delivery_Lead { get; set; }
        public DateTime Start_Date { get; set; }
        public DateTime End_Date { get; set; }
        public string Client_Location { get; set; }
        public string Project_Code { get; set; }
        public string Project_Description { get; set; }
        public string Client_Name { get; set; }
        public string Business_Unit_AGU { get; set; }
        public string Oracle_Comments { get; set; }
        public string Database_Platform { get; set; }
        public string Programming_Language { get; set; }
        public string Infrastructure_Platform { get; set; }
        public DateTime Modified { get; set; }
        public DateTime Created { get; set; }

    }
}

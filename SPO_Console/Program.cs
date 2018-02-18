using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Security;
using Excel = Microsoft.Office.Interop.Excel;

namespace SPO_Console
{
    class Program
    {

        static void Main(string[] args)
        {

            #region Authentication

            string username = "kevinku@microsoft.com";
            string password = "Nexus2013";

            SecureString securepassword = new SecureString();
            foreach (char c in password.ToCharArray())
                securepassword.AppendChar(c);
            var onlineCredentials = new SharePointOnlineCredentials(username, securepassword);

            #endregion
            

            #region Index for columns
            /*
            (1) Number -- SR_x0020_Number -- SR_x0020_Number
            (2) Number -- Age -- Age
            (3) Number -- Total_x0020_SR_x0020_Labor_x0020 -- Total_x0020_SR_x0020_Labor_x0020
            (4) Person or Group -- SR_x0020_Owner -- SR_x0020_Owner
            (5) Single line of text -- Support_x0020_Topic -- Support_x0020_Topic
            (6) Person or Group -- SME_x0020_Reviewer -- SME_x0020_Reviewer
            (7) Single line of text -- Comments -- Comments
            (8) Single line of text -- Opportunities -- Opportunities
            (9) Person or Group -- Last_x0020_week_x0020_SME_x0020_ -- Last_x0020_week_x0020_SME_x0020_
            (10)Single line of text -- Last_x0020_week_x0020_Comments -- Last_x0020_week_x0020_Comments
            (11)Single line of text -- Last_x0020_week_x0020_Opportunit -- Last_x0020_week_x0020_Opportunit

            */
            #endregion

            string siteurl  = "https://microsoft-my.sharepoint.com/personal/kevinku_microsoft_com/subsite";
            string xlsfile  = @"C:\Users\Kevin\Desktop\SRWellnessApril 26.xls";
            string listname = "Cust List";


            #region Opening the Excel file and defining the columns

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(xlsfile);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            List<string> colList = new List<string>();

            for (int j = 1; j <= colCount; j++)
            {

                // Adjusting the Internal name for certain columns which has a longer name

                if (xlRange.Cells[1, j].Value2.ToString() == "Total SR Labor (Mins)")
                    colList.Add("Total_x0020_SR_x0020_Labor_x0020");
                else if (xlRange.Cells[1, j].Value2.ToString() == "Last week SME Reviewer")
                    colList.Add("Last_x0020_week_x0020_SME_x0020_");
                else if (xlRange.Cells[1, j].Value2.ToString() == "Last week Opportunities")
                    colList.Add("Last_x0020_week_x0020_Opportunit");
                else
                    colList.Add(xlRange.Cells[1, j].Value2.ToString().Replace(" ", "_x0020_"));   // Converting all spaces to it's hexadecimal value, _x0020_ 
            }


            Console.WriteLine("Opened the Excel file successfully.....");


            try
            {
                using (ClientContext context = new ClientContext(siteurl))
                {
                    context.Credentials = onlineCredentials;

                    Web oweb = context.Web;
                    context.Load(oweb);
                    context.ExecuteQuery();

                    Console.WriteLine("Successfully connected to the SPO site.....");

                    List olist = oweb.Lists.GetByTitle(listname);
                    ListItemCreationInformation lici = new ListItemCreationInformation();

                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = "<View><RowLimit>200</RowLimit></View>";
                    ListItemCollection allitems = olist.GetItems(camlQuery);

                    context.Load(allitems);
                    context.ExecuteQuery();

                    #region Iterting the field names within the list

                    //FieldCollection flc = olist.Fields;

                    //context.Load(flc);

                    //context.ExecuteQuery();

                    //foreach (Field f in flc)
                    //{
                    //    Console.WriteLine(f.TypeDisplayName + " -- " + f.InternalName + " -- " + f.StaticName);
                    //}

                    #endregion

                    

                    #endregion

                    #region Add a single list item
                    //ListItem olistitem = olist.AddItem(lici);

                    //for (int j = 1; j <= colCount; j++)
                    //{
                    //    if (xlRange.Cells[2, j].Value2 == null)
                    //        continue;
                    //    else if (j == 4 || j == 6 || j == 9)
                    //        olistitem[colList[j - 1]] = oweb.EnsureUser(xlRange.Cells[2, j].Value2.ToString() + "@microsoft.com");
                    //    else
                    //        olistitem[colList[j - 1]] = xlRange.Cells[2, j].Value2.ToString();
                    //}


                    //olistitem.Update();
                    //context.ExecuteQuery();

                    #endregion

                    for (int i = 2; i <= rowCount; i++)
                    {
                        ListItem olistitem = olist.AddItem(lici);
                        ListItem existing = null;

                        if(allitems.Count > 0)
                        {
                            existing = CheckSR(allitems, xlRange.Cells[i, 1].Value2.ToString());
                        }
                        
                        if (existing != null)
                        {
                            existing["Age"] = xlRange.Cells[i, 2].Value2.ToString();
                            existing["Total_x0020_SR_x0020_Labor_x0020"] = xlRange.Cells[i, 3].Value2.ToString();
                            existing["SR_x0020_Owner"] = oweb.EnsureUser(xlRange.Cells[i, 4].Value2.ToString() + "@microsoft.com");
                            existing["Support_x0020_Topic"] = xlRange.Cells[i, 5].Value2.ToString();

                            ClientResult<PrincipalInfo> persons = Utility.ResolvePrincipal(context, oweb, ((Microsoft.SharePoint.Client.FieldLookupValue)(existing["SME_x0020_Reviewer"])).LookupValue, PrincipalType.User, PrincipalSource.All, null, false);
                            context.ExecuteQuery();
                            PrincipalInfo person = persons.Value;

                            existing["Last_x0020_week_x0020_SME_x0020_"] = oweb.EnsureUser(person.Email);
                            existing["Last_x0020_week_x0020_Comments"] = existing["Comments"];
                            existing["Last_x0020_week_x0020_Opportunit"] = existing["Opportunities"];


                            if (xlRange.Cells[i, 6].Value2 != null)
                                existing["SME_x0020_Reviewer"] = oweb.EnsureUser(xlRange.Cells[i, 6].Value2.ToString() + "@microsoft.com");
                            if (xlRange.Cells[i, 7].Value2 != null)
                                existing["Comments"] = xlRange.Cells[i, 7].Value2.ToString();
                            if (xlRange.Cells[i, 8].Value2 != null)
                                existing["Opportunities"] = xlRange.Cells[i, 8].Value2.ToString(); 

                            existing.Update();
                            context.ExecuteQuery();
                            continue;

                        }

                        for (int j = 1; j <= colCount; j++)
                        {
                           if (xlRange.Cells[i, j].Value2 == null)
                              continue;
                           else if (j == 4 || j == 6 || j == 9)
                              olistitem[colList[j - 1]] = oweb.EnsureUser(xlRange.Cells[i, j].Value2.ToString() + "@microsoft.com");
                           else
                              olistitem[colList[j - 1]] = xlRange.Cells[i, j].Value2.ToString();
                        }
                            
                            
                        olistitem.Update();
                        context.ExecuteQuery();
                        lici = null;
                    }                  

                }
            
                Console.WriteLine("Done!!");
                Console.ReadKey();

            }
           catch(Exception ex)
           {
              Console.WriteLine("ERROR:- "+ ex.Message);
              Console.ReadKey();
           }

            #region Closing the Excel connections
            finally
            {
                xlWorkbook.Close(true, null, null);
                xlApp.Quit();


                releaseObject(xlWorksheet);
                releaseObject(xlWorkbook);
                releaseObject(xlApp);

                Console.WriteLine("Excel file closed properly.....");
            }
            #endregion

        }


        #region Checking if it's an existing SR
        public static ListItem CheckSR(ListItemCollection allitems , string SR)
        {
            foreach(ListItem i in allitems)
            {
                if(i["SR_x0020_Number"].ToString() == SR)
                {
                    return i;
                }
            }

            return null;
        }

        #endregion

        #region Closing the Excel file connection methods
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }

            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Unable to release the Object " + ex.ToString());
            }

            finally
            {
                GC.Collect();
            }
        }
        #endregion
    
    }
}

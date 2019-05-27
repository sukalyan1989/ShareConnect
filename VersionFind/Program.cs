using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Security;
namespace VersionFind
{
    class Program
    {
        static void Main(string[] args)
        {
            string connectionstring = @"Data Source=TESTVPS;Initial Catalog=versiondata;Integrated Security=True;Connect Timeout=30";
            List<Vclass> valueList = new List<Vclass>();
                string siteCollectionUrl = "https://cookinsurance.sharepoint.com/sites/112/";
                string userName = "jhart@cookins.us";
                string password = "admin%123%";

                Console.WriteLine("Signing in ..");
                ClientContext ctx = new ClientContext(siteCollectionUrl);
                SecureString secureString = new SecureString();
                password.ToList().ForEach(secureString.AppendChar);

                // Namespace: Microsoft.SharePoint.Client  
                ctx.Credentials = new SharePointOnlineCredentials(userName, secureString);

                // Namespace: Microsoft.SharePoint.Client  
                Site site = ctx.Site;
                Console.WriteLine("Sign in success ..");
                ctx.Load(site);
                ctx.ExecuteQuery();
                Web web = ctx.Web;

                ctx.Load(web, w => w.ServerRelativeUrl, w => w.Lists);

            CheckTableNameAgain:
            Console.WriteLine("Please Enter The SQL Table name : ");
            string tableName = Console.ReadLine();

            //checking if table exists in the database or not
            try
            {
            using (SqlConnection conn = new SqlConnection(connectionstring))
            {        
                conn.Open();
                SqlCommand cmd = new SqlCommand(@"select * from " + tableName);
                cmd.CommandType = CommandType.Text;
                cmd.Connection = conn;
                cmd.ExecuteNonQuery();
                conn.Close();
            }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                goto CheckTableNameAgain;
            }


                ListSearchAgain:
                Console.WriteLine("Enter The Target List Name :");
            string listName = Console.ReadLine();
                
                List list = web.Lists.GetByTitle(listName);

                ctx.Load(list);
                
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View><Query><Where><Geq><FieldRef Name ='ID'/><Value Type='Number'>1</Value></Geq></Where></Query></View>";
                ListItemCollection itemColl = list.GetItems(camlQuery);

                ctx.Load(itemColl);
            try
            {

                ctx.ExecuteQuery();
            }
                catch(Exception e)
            {
                Console.WriteLine(e.Message);
                goto ListSearchAgain;
            }

            if (itemColl.Count == 0)
            {
                Console.WriteLine("Sorry No Data Found");
            }

                foreach (ListItem item in itemColl)

                {

                Console.WriteLine("Handling Item with Id no : "+ item.Id);
                    ListItemVersionCollection itemversioncollection = item.Versions;

                AttachmentCollection itemAttachmentCollection = item.AttachmentFiles;
                ctx.Load(itemAttachmentCollection);
                ctx.ExecuteQuery();


                foreach (Attachment attachment in itemAttachmentCollection)
                {
                    ctx.Load(attachment);
                    ctx.ExecuteQuery();
                    var y = attachment;
                    var file = ctx.Web.GetFileByServerRelativeUrl(y.ServerRelativeUrl);
                    ctx.Load(file);
                    ctx.ExecuteQuery();
                    ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                    ctx.Load(file);
                    ctx.ExecuteQuery();

                    using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                    {
                        if (data != null)
                        {
                           System.IO.Directory.CreateDirectory(Path.Combine(@"C:\",listName, item.Id.ToString()));
                            var fileName = Path.Combine(@"C:\",listName, item.Id.ToString(), file.Name);
                            using (var fileStream = System.IO.File.Create(fileName))
                            {
                                data.Value.CopyTo(fileStream);
                            }
                                data.Value.CopyTo(mStream);
                            byte[] imageArray = mStream.ToArray();
                            string b64String = Convert.ToBase64String(imageArray);
                           
                        }
                    }


                }



                // program to collect version information
                ctx.Load(itemversioncollection);

                    ctx.ExecuteQuery();

                    for (int iVersionCount = 0; iVersionCount < itemversioncollection.Count; iVersionCount++)

                    {

                        ListItemVersion version = itemversioncollection[iVersionCount];

                    var x = version.FileVersion;
                        Vclass v = new Vclass();
                        v.VId = Convert.ToInt32(item.Id);
                    try
                    {
                        v.Title = version.FieldValues["Title"].ToString();

                    }
                    catch
                    {
                                           
                        v.Title = "None";
                        
                    }
                    try {

                        v.Description = version.FieldValues["Body"].ToString();
                    }
                    catch {

                        v.Description = "None";
                    }
                    try {
                        v.StartDate = version.FieldValues["StartDate"].ToString();
                    }
                    catch {
                        v.StartDate = "Not Available";
                    }
                    try
                    {
                        v.DueDate = version.FieldValues["DueDate"].ToString();
                    }
                    catch
                    {
                        v.DueDate = "Not Available";
                    }

                    //Assigned To
                    try
                    {

                        FieldUserValue[] tempATEMail = (FieldUserValue[]) version.FieldValues["AssignedTo"];
                        v.AssignedToEmail = tempATEMail[0].Email.ToString();
                   }
                    catch
                    {
                    v.AssignedToEmail = "None";
                    }
                    try
                    {
                        FieldUserValue[] tempATName = (FieldUserValue[])version.FieldValues["AssignedTo"];
                        v.AssginedToName = tempATName[0].LookupValue.ToString();
                    }
                    catch
                    {
                        v.AssginedToName = "None";
                    }
                    try
                    {
                      v.AppliesTo= version.FieldValues["AppliesTO"].ToString();
                      
                    }
                    catch
                    {
                        v.AppliesTo = "None";

                    }

                    //Follow Up ..
                    try
                    {
                        v.FollowUp = version.FieldValues["FollowUp"].ToString();
                    }
                    catch
                    {
                        v.FollowUp = "None";
                    }


                    //ClientID Look Up value
                    try
                    {

                        FieldLookupValue tempClientId = (FieldLookupValue)version.FieldValues["Client"];
                        v.ClientId = tempClientId.LookupId.ToString();
                    }
                    catch
                    {
                        v.ClientId = "None";
                    }
                    //ClientName Look Up Value
                    try
                    {

                        FieldLookupValue tempClientName = (FieldLookupValue)version.FieldValues["Client"];
                        v.ClientName = tempClientName.LookupValue.ToString();
                    }
                    catch
                    {
                        v.ClientName = "None";
                    }
                    try
                    {
                        v.Status = version.FieldValues["Status"].ToString();
                    }
                    catch
                    {
                        v.Status = "None";
                    }
                    //percent complete

                    try
                    {
                        v.PercentComplete = version.FieldValues["PercentComplete"].ToString();
                    }
                    catch
                    {
                        v.PercentComplete = "None";
                    }
                    //Created Date
                    try
                    {
                        v.CreatedDate = version.FieldValues["Created_x0020_Date"].ToString();
                    }
                    catch
                    {
                        v.CreatedDate = "None";
                    }
                    //Modified Date

                    try
                    {
                        v.ModifiedDate = version.FieldValues["Last_x0020_Date"].ToString();
                    }
                    catch
                    {
                        v.ModifiedDate = "None";
                    }
                    //Priority
                    try
                    {
                        v.Priority = version.FieldValues["Priority"].ToString();
                    }
                    catch
                    {
                        v.Priority = "None";
                    }
                    //Related Items
                    try
                    {
                        v.RelatedItems = version.FieldValues["RelatedItems"].ToString();
                    }
                    catch
                    {
                        v.RelatedItems = "None";
                    }
                    //Modified Name
                    try
                    {
                        FieldUserValue tempModName = (FieldUserValue)version.FieldValues["Editor"];
                        v.ModifiedName = tempModName.LookupValue.ToString();
                    }
                    catch
                    {
                        v.ModifiedName = "None";
                    }

                    //Modified Email
                    try
                    {
                        FieldUserValue tempModEmail = (FieldUserValue)version.FieldValues["Editor"];
                        v.ModifiedEmail = tempModEmail.Email.ToString();
                    }
                    catch
                    {
                        v.ModifiedEmail = "None";
                    }
                    //Created Email
                    try
                    {
                        FieldUserValue tempCreatedemail = (FieldUserValue)version.FieldValues["Editor"];
                        v.CreatedEmail = tempCreatedemail.Email.ToString();
                    }
                    catch
                    {
                        v.CreatedEmail = "None";
                    }
                    v.VersionLabel = version.VersionLabel;
                        valueList.Add(v);
                    }
                


                
            }
            // string connectionstring = configurationmanager.connectionstrings["sharepoint"].connectionstring;
            

            using (SqlConnection conn = new SqlConnection(connectionstring))
            {
                conn.Open();

                SqlCommand cmd =
                    new SqlCommand(
                        "insert into "+tableName+"(VID,Title,Description,VersionLabel,StartDate,DueDate,PercentageComplete,AppliesTo,AssignedToName,AssignedToEmail,ClientName,ClientID,Completed,Created,FollowUp,Modified,Priority,RelatedItems,Status,CreatedBy,ModifiedBy) " +
                        " values (@param1,@param2,@param3,@param4,@param5,@param6,@param7,@param8,@param9,@param10,@param11,@param12,@param13,@param14,@param15,@param16,@param17,@param18,@param19,@param20,@param21)");
                cmd.CommandType = CommandType.Text;
                cmd.Connection = conn;
                cmd.Parameters.Add("@param1", DbType.String);
                cmd.Parameters.Add("@param2", DbType.String);
                cmd.Parameters.Add("@param3", DbType.String);
                cmd.Parameters.Add("@param4", DbType.String);
                cmd.Parameters.Add("@param5", DbType.String);
                cmd.Parameters.Add("@param6", DbType.String);
                cmd.Parameters.Add("@param7", DbType.String);
                cmd.Parameters.Add("@param8", DbType.String);
                cmd.Parameters.Add("@param9", DbType.String);
                cmd.Parameters.Add("@param10", DbType.String);
                cmd.Parameters.Add("@param11", DbType.String);
                cmd.Parameters.Add("@param12", DbType.String);
                cmd.Parameters.Add("@param13", DbType.String);
                cmd.Parameters.Add("@param14", DbType.String);
                cmd.Parameters.Add("@param15", DbType.String);
                cmd.Parameters.Add("@param16", DbType.String);
                cmd.Parameters.Add("@param17", DbType.String);
                cmd.Parameters.Add("@param18", DbType.String);
                cmd.Parameters.Add("@param19", DbType.String);
                cmd.Parameters.Add("@param20", DbType.String);
                cmd.Parameters.Add("@param21", DbType.String);
                



                foreach (var item in valueList)
                {
                    cmd.Parameters[0].Value = item.VId;
                    cmd.Parameters[1].Value = item.Title;
                    cmd.Parameters[2].Value = item.Description;
                    cmd.Parameters[3].Value = item.VersionLabel;
                    cmd.Parameters[4].Value = item.StartDate;
                    cmd.Parameters[5].Value = item.DueDate;
                    cmd.Parameters[6].Value = item.PercentComplete;
                    cmd.Parameters[7].Value = item.AppliesTo;
                    cmd.Parameters[8].Value = item.AssginedToName;
                    cmd.Parameters[9].Value = item.AssignedToEmail;
                    cmd.Parameters[10].Value = item.ClientName;
                    cmd.Parameters[11].Value = item.ClientId;
                    cmd.Parameters[12].Value = item.PercentComplete;
                    cmd.Parameters[13].Value = item.CreatedDate;
                    cmd.Parameters[14].Value = item.FollowUp;
                    cmd.Parameters[15].Value = item.ModifiedDate;
                    cmd.Parameters[16].Value = item.Priority;
                    cmd.Parameters[17].Value = item.RelatedItems;
                    cmd.Parameters[18].Value = item.Status;
                    cmd.Parameters[19].Value = item.CreatedEmail;
                    cmd.Parameters[20].Value = item.ModifiedEmail;
                    
                   
               
                    
                    cmd.ExecuteNonQuery();
                }

                conn.Close();
            }
            Console.WriteLine("done... Press Any Key to close");
        }
        
    }
}

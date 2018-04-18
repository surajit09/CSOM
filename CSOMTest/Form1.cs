using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Security;

namespace CSOMTest
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Load_Click(object sender, EventArgs e)
        {
            System.Net.HttpWebRequest httpWebRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create("https://thethread.carpetright.co.uk/Facilities/");
            httpWebRequest.UseDefaultCredentials = true;
            using (var context = new ClientContext("https://thethread.carpetright.co.uk/Facilities/"))
            {
                var web = context.Web;
                context.Load(web);
                context.Load(web.Lists);
                context.ExecuteQuery();
                ResultsListBox.Items.Add(web.Title);
                ResultsListBox.Items.Add(web.Lists.Count);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var context = new ClientContext("https://thethread.carpetright.co.uk/Facilities/"))
            {
                var web = context.Web;
                var query = from list in web.Lists
                            where list.Title== "Manager's Weekly Check"
                            && list.ItemCount > 0
                            select list;
                
                var lists=context.LoadQuery(query);
                context.ExecuteQuery();
                //ResultsListBox.Items.Add(web.Title);
                ResultsListBox.Items.Add(lists.Count());
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var Url = "https://thethread-qa.carpetright.co.uk/sites/AppTest/_api/web";

            var client = new WebClient();

            client.UseDefaultCredentials = true;

            var xml = client.DownloadString(Url);

            var doc = XDocument.Parse(xml);



        }

        private void button3_Click(object sender, EventArgs e)
        {
            var Url = "https://thethread-qa.carpetright.co.uk/sites/AppTest/_api/web";

            var client = new WebClient();

            client.UseDefaultCredentials = true;
            client.Headers[HttpRequestHeader.Accept] = "application/json;odata=verbose";
            var json = client.DownloadString(Url);

            var ser = new JavaScriptSerializer();
            dynamic item = ser.Deserialize<object>(json);
            ResultsListBox.Items.Add(item["d"]["Title"]);


        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (var context = new ClientContext("https://thethread.carpetright.co.uk/Facilities/"))
            {
                var web = context.Web;
                context.Load(web,w=>w.Title, w => w.Description);
                //context.Load(web.Lists, l=>l.Count);
                context.ExecuteQuery();
                ResultsListBox.Items.Add(web.Title);
                ResultsListBox.Items.Add(web.Description);
                //ResultsListBox.Items.Add(web.Lists.Count);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
           
            using (var context = new ClientContext("https://thethread.carpetright.co.uk/Facilities/"))
            {
                var web = context.Web;
                context.Load(web, w => w.Title, w => w.Description);
                var lists = web.Lists;
                context.Load(lists, lc=>lc.Include(
                    l=>l.Title,
                    l=>l.Fields.Include(
                        f=>f.Title
                        
                        )
                    
                    ));
                context.ExecuteQuery();
                ResultsListBox.Items.Add(web.Title);
                foreach (var list in lists)
                {

                    ResultsListBox.Items.Add(list.Title);
                    foreach(var field in list.Fields)
                    {
                        ResultsListBox.Items.Add("\t"+field.Title);
                    }

                }



                //ResultsListBox.Items.Add(web.Lists.Count);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.Net.HttpWebRequest httpWebRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create("https://thethread.carpetright.co.uk/Facilities/");
            httpWebRequest.UseDefaultCredentials = true;
            using (var context = new ClientContext("https://thethread.carpetright.co.uk/Facilities/"))
            {
                var web = context.Web;
                context.Load(web, w => w.Title, w => w.Description);
                var list = web.Lists.GetByTitle("Manager's Weekly Check");
                var query = new CamlQuery();
                query.ViewXml = "<View/>";
                var items = list.GetItems(query);
                context.Load(list, l=>l.Title);
                context.Load(items, c=>c.Include(li=>li["Title"]));
                context.ExecuteQuery();
                ResultsListBox.Items.Add(web.Title);
                ResultsListBox.Items.Add(web.Description);
                foreach(var item in items)
                {
                    ResultsListBox.Items.Add(item["Title"]);
                }

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            using (var context = new ClientContext("https://thethread-qa.carpetright.co.uk/Facilities/"))
            {
                try
                {
                    var web = context.Web;
                    List list = null;
                    var scope = new ExceptionHandlingScope(context);
                    using (scope.StartScope())
                    {
                        using (scope.StartTry())
                        {
                            list = web.Lists.GetByTitle("Test Tasks");
                            context.Load(list);
                        }
                        using (scope.StartCatch())
                        {
                            var lci = new ListCreationInformation();
                            lci.Title = "Test Tasks";
                            lci.QuickLaunchOption = QuickLaunchOptions.On;
                            lci.TemplateType = (int)ListTemplateType.Tasks;
                             list = web.Lists.Add(lci);
                        }

                        using (scope.StartFinally())
                        {
                            list = web.Lists.GetByTitle("Test Tasks");
                            context.Load(list);
                        }
                    }


                    //context.Load(web);

                    context.ExecuteQuery();

                    var status = scope.HasException ? "Created" : "Loaded";
                    ResultsListBox.Items.Add(list.Title +status );
                    
                }
                catch(Exception ex)
                {
                    ResultsListBox.Items.Add(ex.Message);
                }

               
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            using (var context = new ClientContext("https://thethread-qa.carpetright.co.uk/Facilities/"))
            {
                try
                {
                    var web = context.Web;
                    List list = null;
                    var scope = new ExceptionHandlingScope(context);
                    using (scope.StartScope())
                    {
                        using (scope.StartTry())
                        {
                            list = web.Lists.GetByTitle("Test Tasks");
                            var ici = new ListItemCreationInformation();
                            var item = list.AddItem(ici);
                            item["Title"] = "Sample Task";
                            item["AssignedTo"] = web.CurrentUser;
                            item["DueDate"] = DateTime.Now.AddDays(7);
                            item.Update();
                            //context.Load(list);
                        }
                        using (scope.StartCatch())
                        {
                            var lci = new ListCreationInformation();
                            lci.Title = "Test Tasks";
                            lci.QuickLaunchOption = QuickLaunchOptions.On;
                            lci.TemplateType = (int)ListTemplateType.Tasks;
                            list = web.Lists.Add(lci);
                        }

                        using (scope.StartFinally())
                        {
                           // list = web.Lists.GetByTitle("Test Tasks");
                           // context.Load(list);
                        }
                    }


                    //context.Load(web);

                    context.ExecuteQuery();

                    var status = scope.HasException ? "Not Created" : "Created";
                    ResultsListBox.Items.Add("Item " + status);

                }
                catch (Exception ex)
                {
                    ResultsListBox.Items.Add(ex.Message);
                }


            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            using (var context = new ClientContext("https://thethread-qa.carpetright.co.uk/Facilities/"))
            {
                try
                {
                    var web = context.Web;
                    List list = null;
                    ListItemCollection items = null;
                    var scope = new ExceptionHandlingScope(context);
                    using (scope.StartScope())
                    {
                        using (scope.StartTry())
                        {
                            list = web.Lists.GetByTitle("Test Tasks");
                            var query = new CamlQuery();
                            query.ViewXml = "<View><RowLimit>1</RowLimit></View>";
                             items = list.GetItems(query);
                            
                            
                            context.Load(items);
                        }
                        using (scope.StartCatch())
                        {
                            var lci = new ListCreationInformation();
                            lci.Title = "Test Tasks";
                            lci.QuickLaunchOption = QuickLaunchOptions.On;
                            lci.TemplateType = (int)ListTemplateType.Tasks;
                            list = web.Lists.Add(lci);
                        }

                        using (scope.StartFinally())
                        {
                            // list = web.Lists.GetByTitle("Test Tasks");
                            // context.Load(list);
                        }
                    }


                    //context.Load(web);

                    context.ExecuteQuery();
                    var item = items.FirstOrDefault();
                    if(item != null)
                    {
                        item["Status"] = "In Progress";
                        item["PercentComplete"] = 0.1;
                        item.Update();
                    }
                    context.ExecuteQuery();
                   // var status = scope.HasException ? "Not Created" : "Created";
                    ResultsListBox.Items.Add("Item updated");

                }
                catch (Exception ex)
                {
                    ResultsListBox.Items.Add(ex.Message);
                }


            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            var siteUrl = "https://carpetright.sharepoint.com/TestDevFrameWork/";
            var loginName = "hosme@carpetright.co.uk";
            var password = "MewMew0903";


            var securePassword = new SecureString();
            password.ToCharArray().ToList().ForEach(c => securePassword.AppendChar(c));

            using(var context=new ClientContext(siteUrl))
            {
                context.Credentials = new SharePointOnlineCredentials(loginName, securePassword);
                var web = context.Web;
                var list = web.Lists.GetByTitle("ColumnFormatter");

                var query = new CamlQuery();
                query.ViewXml = "<View />";
                 var items = list.GetItems(query);

                context.Load(items, c => c.Include(li => li.Client_Title));
                context.ExecuteQuery();
                foreach(var item in items)
                {
                    ResultsListBox.Items.Add(item.Client_Title);
                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            using (var context = new ClientContext("https://thethread-qa.carpetright.co.uk/Facilities/"))
            {
                try
                {
                    var web = context.Web;
                    List list = null;
                    
                       var lci = new ListCreationInformation();
                            lci.Title = "Project Documents";
                            lci.QuickLaunchOption = QuickLaunchOptions.On;
                            lci.TemplateType = (int)ListTemplateType.DocumentLibrary;
                            list = web.Lists.Add(lci);

                            FieldType fieldType = FieldType.Number;
                            OfficeDevPnP.Core.Entities.FieldCreationInformation newFieldInfo = new OfficeDevPnP.Core.Entities.FieldCreationInformation(fieldType);
                            newFieldInfo.DisplayName = "Year";
                            newFieldInfo.AddToDefaultView = true;
                            newFieldInfo.InternalName = "Year";
                            newFieldInfo.Id = Guid.NewGuid();

                            list.CreateField(newFieldInfo,false);

                    context.ExecuteQuery();
                    var status ="Field Created";
                    ResultsListBox.Items.Add(list.Title + status);


                }


                  
                catch (Exception ex)
                {
                    ResultsListBox.Items.Add(ex.Message);
                }


            }


        }
    }
}

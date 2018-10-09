using System;
using System.Linq;
using System.Security;
using Microsoft.SharePoint.Client;

namespace csonConsoleApplication
{
    partial class Program
    {
        public static void Main(string[] args)
        {
            string userName = "dinesh.bhatt@acuvate.com";
            Console.WriteLine("Enter your password.");
            SecureString password = GetPassword();
        
            using (var clientContext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/IJOX"))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(userName, password);
                Web rootWeb = clientContext.Web;
                clientContext.Load(rootWeb, w => w.Title, w => w.Description, w => w.Url);
                RemoveUser(clientContext);
                //AddList(clientContext);
                //DeleteList(clientContext);
                //CreateDocumentLiberary(clientContext);
                //List myList = rootWeb.Lists.GetByTitle("EmployeeContact");
                //clientContext.Load(myList);
                //UploadFile(clientContext);
                //ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                //ListItem ListItem = myList.AddItem(itemCreateInfo);
                //ListItem["Title"] = "Dinesh";
                //ListItem["First Name"] = "Chandra";
                //ListItem["Email Address"] = "chandra.bhatt@gamil.com";
                //myList.Update();
                //myList.Fields.AddFieldAsXml(@"<Field Name='Age' DisplayName='Age' type='Number' Required='FALSE'></Field>", true, AddFieldOptions.DefaultValue);
                //myList.Update();
                clientContext.ExecuteQuery();
                Console.WriteLine("Title: " + rootWeb.Title + "; URL: " + rootWeb.Url);
                Console.ReadLine();
            }
            
        }
        private static SecureString GetPassword()
        {
            ConsoleKeyInfo info;
            //Get the user's password as a SecureString  
            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }
        public static void CreateSubSite( ClientContext context )
        {
            WebCreationInformation webCreationInformation = new WebCreationInformation();
            Console.WriteLine("enter the url");
            webCreationInformation.Url = Console.ReadLine().Trim().Replace(" ", "");
            Console.WriteLine("enter the site name");
            webCreationInformation.Title = Console.ReadLine();
            context.Web.Webs.Add(webCreationInformation);
            try
            {
                context.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.WriteLine("new error", e);
                throw e;
            }
          
        }
        public void AddColumn(ClientContext context)
        {
            List list = context.Web.Lists.GetByTitle("StudentDetails");
            context.Load(list);
            Field fields = list.Fields.GetByTitle("columnName");
            fields.DeleteObject();
            
            context.ExecuteQuery();
        }
        public static void UploadFile(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("Document Libarary");
            FileCreationInformation fileCreationInformation = new FileCreationInformation();
            fileCreationInformation.Content = System.IO.File.ReadAllBytes(@"C:\Users\dinesh.bhatt\source\repos\SharePointCSOM\csonConsoleApplication\requiremnet.txt");
            fileCreationInformation.Url = @"Document Libarary\requiremnet.txt";
            fileCreationInformation.Overwrite = true;
            File fileToUpload = myList.RootFolder.Files.Add(fileCreationInformation);
            ctx.Load(fileToUpload);
         
            ctx.ExecuteQuery();
           
            
        }
        public static void AddList(ClientContext clientContext)
        {
            Web newWeb = clientContext.Web;
            clientContext.Load(newWeb);
            ListCreationInformation listCreationInformation = new ListCreationInformation();
            listCreationInformation.Title = "Folder1";
            listCreationInformation.Description = "new list to upload the files in it";
            listCreationInformation.TemplateType = (int)ListTemplateType.Announcements;
            List newList = newWeb.Lists.Add(listCreationInformation);
            
            newList.Update();
            clientContext.ExecuteQuery();
        }
        public static void AddUser(ClientContext context)
        {
            Web web = context.Web;
            User user = web.EnsureUser("arvind.torvi@acuvate.com");
            Group group = web.SiteGroups.GetByName("IJOX Owners");
            group.Users.AddUser(user);
        }
        public static void RemoveUser(ClientContext client)
        {
            Web web = client.Web;
            User user = web.EnsureUser("IJOX Owners");
            Group testingOwnersGroup = web.SiteGroups.GetByName("IJOX Owners");
            
            UserCollection userCollection = testingOwnersGroup.Users;

            userCollection.Remove(user);
        }
        public static void DeleteList(ClientContext client)
        {
            List list = client.Web.Lists.GetByTitle("folder1");
            list.DeleteObject();
            client.ExecuteQuery();
        }
        public static void CreateDocumentLiberary(ClientContext context)
        {
            Web newWeb = context.Web;
            context.Load(newWeb);
            ListCreationInformation listCreationInformation = new ListCreationInformation();
            listCreationInformation.Title = "Document Libarary";
            listCreationInformation.Description = "new list to upload the files in it";
            listCreationInformation.TemplateType = (int)ListTemplateType.DocumentLibrary;
            List newList = newWeb.Lists.Add(listCreationInformation);
            newList.Update();
            context.ExecuteQuery();
        }
    }
}


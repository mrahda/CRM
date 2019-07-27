using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using System.Net;
using System.IO;
using Microsoft.Xrm.Sdk.Query;
using Aspose.Words;

namespace CRM
{
    class Program
    {
        static void Main(string[] args)
        {
            //Get Organisation 
            string organizationUrl = "http://tashilat";
            if (!organizationUrl.EndsWith("/")) organizationUrl += "/";
            Uri orgUri = new Uri(string.Format("{0}{1}/XRMServices/2011/Organization.svc", organizationUrl, "NSFUND"));
            ClientCredentials credentials = new ClientCredentials();
            credentials.Windows.ClientCredential = new NetworkCredential("crmtest2", "!9358891444!@", "");
            OrganizationServiceProxy orgService = new OrganizationServiceProxy(orgUri, null, credentials, null);
            orgService.EnableProxyTypes();

            //CreateAnnotaion(orgService);
            //ReadNote(orgService);
            //TestCriteri(orgService);
            CreatePDF();

            Console.ReadKey();
        }

        private static void CreatePDF()
        {
            LicenseHelper.ModifyInMemory.ActivateMemoryPatching();
            Document doc = new Document(@"d:\sefid\کاردکس جدید تسهیلات  -محاسبات .docx");
            doc.Save(@"d:\sefid\123.pdf", SaveFormat.Pdf);
            Console.Write("OK");
        }

        private static void TestCriteri(OrganizationServiceProxy orgService)
        {
            QueryExpression qry = new QueryExpression("account");
            qry.ColumnSet = new ColumnSet(true);
            //qry.Criteria.AddCondition("sfd_national_code", ConditionOperator.Equal, "8765432345678");
            //qry.Criteria.AddCondition("sfd_economical_code", ConditionOperator.Equal, "12345676543");

            FilterExpression filter = new FilterExpression(LogicalOperator.Or);
            filter.AddCondition("sfd_economical_code", ConditionOperator.Equal, "12345676543");
            qry.Criteria.AddFilter(filter);

            EntityCollection lst = orgService.RetrieveMultiple(qry);
            qry.AddOrder("modifiedon", OrderType.Descending);
            Console.Write("count:" + lst.Entities.Count.ToString());
        }

        private static void ReadNote(OrganizationServiceProxy orgService)
        {
            Entity OrgNote = orgService.Retrieve("annotation", Guid.Parse("{B911F299-8EAC-E911-98F2-005056AE3285}"), new Microsoft.Xrm.Sdk.Query.ColumnSet("subject","notetext","filename","documentbody","objectid","mimetype"));
            //if (OrgNote.Attributes.Keys.Contains("subject"))
                Console.WriteLine(OrgNote["mimetype"]);
        }

        static void CreateAnnotaion(OrganizationServiceProxy orgService)
        {
            Entity NewNote = new Entity("annotation");
            NewNote["filename"] = @"C:\Windows\TEMP" + @"\" + "b1b511d5-032c-4b27-9d12-886169d38fc1" + ".pdf"; // OrgNote["filename"] + ".pdf";
            NewNote["documentbody"] = Convert.ToBase64String(File.ReadAllBytes(@"C:\Windows\TEMP\" + @"\" + "b1b511d5-032c-4b27-9d12-886169d38fc1" + ".pdf"));
            NewNote["subject"] = "new subject";
            NewNote["notetext"] = "created with code";
            NewNote["objectid"] = new EntityReference("account", Guid.Parse("{FE9ECAE9-9570-E911-80F3-000C29364209}"));
            Guid attachmentId = orgService.Create(NewNote);
            Console.WriteLine(attachmentId.ToString());

        }
    }
}

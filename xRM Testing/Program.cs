using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.ServiceModel;
using System.ServiceModel.Description;
using System.IO;
using Microsoft.Win32;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net;
using Newtonsoft.Json;

namespace xRM_Testing
{
    public class CredentialProvider
    {
        public static string CRMUser
        {
            get
            {
                return ConfigurationManager.AppSettings["CRMUser"];
            }
        }
        public static string CRMPwd
        {
            get
            {
                return ConfigurationManager.AppSettings["CRMPwd"];
            }
        }
        public static ClientCredentials CRMCreds
        {
            get
            {
                ClientCredentials cre = new ClientCredentials();
                cre.UserName.UserName = CredentialProvider.CRMUser;
                cre.UserName.Password = CredentialProvider.CRMPwd;
                return cre;
            }
        }
    }
    public class Lead
    {
        public string Subject { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string NoteSubject { get; set; }
        public string FileName { get; set; }
        public string MimeType { get; set; }
        public byte[] Data { get; set; }
    }
    class Program
    {
        public static string GetContentType(string fname)
        {
            string contentType = "application/octet-stream";

            try
            {
                RegistryKey classes = Registry.ClassesRoot;
                RegistryKey fileClass = classes.OpenSubKey(Path.GetExtension(fname));
                contentType = fileClass.GetValue("Content Type").ToString();
            }
            catch { }

            return contentType;
        }
        public static void createLead(string subject, string firstName, string lastName)
        {
            Uri serviceUri = new Uri(ConfigurationManager.AppSettings["CRMUri"]);
            OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, CredentialProvider.CRMCreds, null);
            proxy.EnableProxyTypes();
            IOrganizationService service = proxy;

            using (XrmServiceContext orgContext = new XrmServiceContext(service))
            {
                Entity lead = new Entity("lead");
                lead.Attributes["subject"] = subject;
                lead.Attributes["firstname"] = firstName;
                lead.Attributes["lastname"] = lastName;
                Guid leadId = service.Create(lead);

                EntityReference entRef = new EntityReference("lead", leadId);

                string fileName = "C:\\Users\\Dalil\\Documents\\Laetitia Casta CV.docx";
                FileStream stream = File.OpenRead(fileName);
                byte[] byteData = new byte[stream.Length];
                stream.Read(byteData, 0, byteData.Length);
                stream.Close();

                Annotation note = new Annotation();
                note.Subject = "Test";
                note.FileName = fileName;
                note.DocumentBody = Convert.ToBase64String(byteData);
                note.MimeType = GetContentType(fileName);
                note.ObjectId = entRef;
                service.Create(note);
            }
        }
        static void Main(string[] args)
        {
            string fileName = "C:\\Users\\dalild\\Documents\\Laetitia CV.docx";
            string responseMessage = null;
            FileStream stream = File.OpenRead(fileName);
            byte[] byteData = new byte[stream.Length];

            JsonSerializer serializer = new Newtonsoft.Json.JsonSerializer();
            Lead lead = new Lead();
            lead.Subject = "Candidate";
            lead.FirstName = "Laetitia";
            lead.LastName = "Casta";
            lead.NoteSubject = "CV";
            lead.FileName = "Laetitia CV.docx";
            lead.MimeType = GetContentType(fileName);
            lead.Data = byteData;

            var json = JsonConvert.SerializeObject(lead);
            byte[] byteArray = Encoding.UTF8.GetBytes(json);

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://192.168.1.28/xRMConnector/api/Lead");
            request.ContentType = "application/json";
            request.Method = "POST";
            request.ContentLength = byteArray.Length;

            Stream dataStream = request.GetRequestStream();
            dataStream.Write(byteArray, 0, byteArray.Length);
            dataStream.Close();

            HttpWebResponse response = (HttpWebResponse)(request.GetResponse());

            if (response.StatusCode == HttpStatusCode.OK)
            {
                Stream responseStream = response.GetResponseStream();

                if (responseStream != null)
                {
                    var reader = new StreamReader(responseStream);
                    responseMessage = reader.ReadToEnd();
                }
            }
            else
            {
                responseMessage = response.StatusDescription;
            }

            Console.WriteLine(responseMessage);
        }
    }
}

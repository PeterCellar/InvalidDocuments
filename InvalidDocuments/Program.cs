using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.IO;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using Microsoft.Xrm.Sdk.Metadata;
using System.IO.Compression;
using System.Collections.Generic;
using System.Threading;
using System.Text.RegularExpressions;
using Microsoft.TeamFoundation.Common;
using System.ServiceModel;
using Microsoft.TeamFoundation.Client.CommandLine;
using Microsoft.TeamFoundation.WorkItemTracking.Process.WebApi.Models;
using Microsoft.Xrm.Tooling.Connector;
using System.Diagnostics;

namespace invalidDocuments
{
    public class Program
    {
        /* 
         * Connects to CRM On-Premises 
         */
        public static void Run(string username, string pswd)
        {
            // Provides an object representation of uniform resource identifier
            Uri homeRealUri = null;
            // Enables to configure client and service credentials as well as credential authentication
            ClientCredentials credentials;
            Uri organizationUri;
            IOrganizationService service;
            OrganizationServiceProxy serviceProxy;
            credentials = new ClientCredentials();

            // CRM On-Premise
            credentials.Windows.ClientCredential = new NetworkCredential(username, pswd);

            // CRM online
            //credentials.UserName.UserName = "ACMARK\\cellar";
            //credentials.UserName.Password = "EbHNhYCaG0BD";

            // Using https for SSL configuration of CRM
            string orgUrl = "https://dev.acmark.eu:5555/ACMARK/XRMServices/2011/Organization.svc";

            // Checking if url contains https
            if (!string.IsNullOrEmpty(orgUrl) && orgUrl.Contains("https"))
            {
                Console.WriteLine("Setting Https...");
                ServicePointManager.ServerCertificateValidationCallback =
                    delegate (object s,
                    X509Certificate certificate,
                    X509Chain chain,
                    SslPolicyErrors sslPolicyErrors)
                    { return true; };
            }

            organizationUri = new Uri(orgUrl);
            using (serviceProxy = new OrganizationServiceProxy(organizationUri, homeRealUri, credentials, null))
            {
                try
                {
                    // serviceProxy.CallerId = Guid.Parse(input: "ACMARK\\cellar");
                    serviceProxy.ServiceConfiguration
                                .CurrentServiceEndpoint
                                .Behaviors
                                .Add(new ProxyTypesBehavior());

                    service = serviceProxy;
                    EntityCollection retrievedRecords = RetrieveRecords(service);
                    Console.WriteLine("Service established. Importing/updating records...");
                    ReadCsvFile("ops_vse.csv", service, retrievedRecords);
                    ReadCsvFile("op_vse.csv", service, retrievedRecords);
                    ReadCsvFile("cd_vse.csv", service, retrievedRecords);
                }
                catch (InvalidOperationException e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }


        /// <summary>
        /// Reads data from downloaded file line by line
        /// Decides if to update an existing record or to import a new record
        /// </summary>
        /// <param name="filePath">Path to downloaded file</param>
        /// <param name="service">Provides programmatic access to the metadata and data for an organization</param>>
        /// <param name="retrievedRecords">Collcetion of records to update</param>
        public static void ReadCsvFile(string filePath, IOrganizationService service, EntityCollection retrievedRecords)
        {
            string docnum = string.Empty, serie = string.Empty, invaliddate = string.Empty;
            Regex checkSerie = new Regex(@"\b[A-Z]{2}[0-9]{0,2}\b");
            Regex checkDate = new Regex(@"\b\d{2}\.\d{0,2}\.\d{4}\b");
            Regex semicolonDelimiter = new Regex(";");
            Guid id;

            using (StreamReader reader = new StreamReader(filePath))
            {
                string value = reader.ReadLine();
                // Loop through every line of a file
                while (value != null)
                {
                    // Split by semicolons
                    string[] semicolonSubstrings = semicolonDelimiter.Split(value);
                    // If first substring is not null or empty then it contains a number of a document
                    if (!semicolonSubstrings[0].IsNullOrEmpty())
                        docnum = semicolonSubstrings[0];

                    // If second substring is not null or empty then it containts serie or invalidDate
                    if (!semicolonSubstrings[1].IsNullOrEmpty())
                    {
                        // Check if substring contains serie of document
                        Match matchedSerie = checkSerie.Match(semicolonSubstrings[1]);
                        if (matchedSerie.Length != 0)
                            serie = semicolonSubstrings[1];

                        // Check if substring contains invalidation date
                        Match matchedDate = checkDate.Match(semicolonSubstrings[1]);
                        if (matchedDate.Length != 0)
                            invaliddate = semicolonSubstrings[1];

                        // If third substring is not null or empty it contains invalidation date
                        if (semicolonSubstrings.Length > 2 && !semicolonSubstrings[2].IsNullOrEmpty())
                        {
                            matchedDate = checkDate.Match(semicolonSubstrings[2]);
                            if (matchedDate.Length != 0)
                                invaliddate = semicolonSubstrings[2];
                        }
                    }


                    // Check if there are already any records in crm
                    if (retrievedRecords.TotalRecordCount != 0)
                        // If there is already a record with specific document number then update record 
                        if ((id = checkForRecordWithSameDocNum(retrievedRecords, docnum)) != Guid.Empty)
                            updateRecord(service, id, invaliddate, docnum, serie);
                        // If there is not a record with specific document number create new record
                        else
                            ImportData(service, docnum, serie, invaliddate);
                    // If there are no records then make first import of the records
                    else
                        ImportData(service, docnum, serie, invaliddate);

                    docnum = string.Empty;
                    invaliddate = string.Empty;
                    serie = string.Empty;

                    value = reader.ReadLine();
                }
            }

        }

        /// <summary>
        /// Updates record with specific Id
        /// </summary>
        /// <param name="service">Provides programmatic access to the metadata and data for an organization</param>
        /// <param name="id">Id of record to be updated</param>
        /// <param name="date">New date to be updated</param>
        /// <param name="num">New document number to be updated</param>
        /// <param name="serie">New serie of document to be updated</param>
        public static void updateRecord(IOrganizationService service, Guid id, string date, string num, string serie)
        {
            Entity updateRecord = new Entity();
            updateRecord.LogicalName = "acm_listinvaliddocument";
            updateRecord.Id = id;
            updateRecord.Attributes["acm_documentnumber"] = num;
            updateRecord.Attributes["acm_batch"] = serie;
            if (date != string.Empty)
                updateRecord.Attributes["acm_invalidationdate"] = DateTime.Parse(date);

            try
            {
                service.Update(updateRecord);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                Console.WriteLine("Message: {0}", ex.Detail.Message);
            }
            catch (System.TimeoutException ex)
            {
                Console.WriteLine("Message: {0}", ex.Message);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            Console.WriteLine("Updated record with number: " + num);
        }

        /// <summary>
        /// Loops through every retrieved record
        /// Checks if there is a record with specific document number
        /// </summary>
        /// <param name="retrievedRecords">Collection of retrieved records where to look for specific document number</param>
        /// <param name="documentNumber">Document number to look for in retrieved records</param>
        /// <returns>Id of a record witch correct document number or empty id</returns>
        public static Guid checkForRecordWithSameDocNum(EntityCollection retrievedRecords, string documentNumber)
        {
            Guid notFound = Guid.Empty;
            // Iterate through all retrieved records
            foreach (Entity entity in retrievedRecords.Entities)
            {
                // Check if document number of retrieved number is equal to a document number of record to be imported
                if (entity.GetAttributeValue<string>("acm_documentnumber").Equals(documentNumber))
                {
                    return entity.GetAttributeValue<Guid>("acm_listinvaliddocumentid");
                }

            }

            return notFound;
        }

        /// <summary>
        /// Imports new records to crm
        /// </summary>
        /// <param name="service">Provides programmatic access to the metadata and data for an organization</param>
        /// <param name="documentNumber">Document number of a record to be imported</param>
        /// <param name="serie">Serie of a record to be imported</param>
        /// <param name="invalidationDate">Invalidation date of a record to be updated</param>
        public static void ImportData(IOrganizationService service, string documentNumber, string serie, string invalidationDate)
        {
            int czechLanguageCode = 1029;
            DateTime? invalidDate = new DateTime();


            Dictionary<string, CrmDataTypeWrapper> maybe = new Dictionary<string, CrmDataTypeWrapper>();
            maybe.Add("OP bez série", new CrmDataTypeWrapper(123, CrmFieldType.Picklist));

            /*PickList picklist = new PickList();
          
             List<string> picklistAttrib = new List<string>();
             picklistAttrib.Add("805210000");
             picklistAttrib.Add("80521000");
             picklistAttrib.Add("805210002");
             picklistAttrib.Add("805210003");

             picklist.Items = picklistAttrib;*/
            OptionSetValueCollection types = new OptionSetValueCollection();
            types.Add(new OptionSetValue(1));
            types.Add(new OptionSetValue(2));

            PicklistAttributeMetadata documentTypePicklist = new PicklistAttributeMetadata()
            {
                SchemaName = "acm_documenttype",
                LogicalName = "acm_documenttype",
                //DisplayName = new Label("Typ dokladu", czechLanguageCode),
                OptionSet = new OptionSetMetadata()
                {
                    //IsGlobal = false,
                    OptionSetType = OptionSetType.Picklist,
                    Options =
                    {
                        new OptionMetadata(new Label("OP bez série", czechLanguageCode), 805210000),
                        new OptionMetadata(new Label("OP se sérií", czechLanguageCode), 805210001),
                        new OptionMetadata(new Label("Cestovní pas - fialový", czechLanguageCode), 805210002),
                        new OptionMetadata(new Label("Cestovní pas - zelený", czechLanguageCode), 805210000)
                    }
                }
            };
            documentTypePicklist.DefaultFormValue = 805210000;


            Entity record = new Entity();
            record.LogicalName = "acm_listinvaliddocument";
            record.Attributes = new AttributeCollection();
            record.Attributes.Add(new KeyValuePair<string, object>("acm_documentnumber", documentNumber));
            record.Attributes.Add(new KeyValuePair<string, object>("acm_batch", serie));

            if (!invalidationDate.IsNullOrEmpty())
                record.Attributes.Add(new KeyValuePair<string, object>("acm_invalidationdate", DateTime.Parse(invalidationDate)));
            else
                record.Attributes.Add(new KeyValuePair<string, object>("acm_invalidationdate", invalidDate = null));
            //Thread.Sleep(1000);
            //record.Attributes["acm_documenttype"] = types;
            //record.Attributes.Add(new KeyValuePair<string, object>("acm_documenttype", new OptionMetadata(new Label("OP bez série", czechLanguageCode), 805210000)));

            try
            {
                Guid recordId = service.Create(record);
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Console.WriteLine("Message: {0}", ex.Detail.Message);
            }
            catch (System.TimeoutException ex)
            {
                Console.WriteLine("Message: {0}", ex.Message);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            Console.WriteLine("Created record with number: " + documentNumber);
        }

        /// <summary>
        /// Retrieves records from CRM
        /// </summary>
        /// <param name="service">Provides programmatic access to the metadata and data for an organization</param>
        /// <returns>A collection of retrieved records</returns>
        public static EntityCollection RetrieveRecords(IOrganizationService service)
        {
            EntityCollection retrievedRecords = new EntityCollection();
            // Query to get existing records in crm
            QueryExpression query = new QueryExpression
            {
                EntityName = "acm_listinvaliddocument",
                ColumnSet = new ColumnSet("acm_documentnumber", "acm_listinvaliddocumentid")
            };

            try
            {
                retrievedRecords = service.RetrieveMultiple(query);
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Console.WriteLine("Message: {0}", ex.Detail.Message);
            }
            catch (System.TimeoutException ex)
            {
                Console.WriteLine("Message: {0}", ex.Message);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            return retrievedRecords;
        }


        /// <summary>
        ///  Downloads .zip files containing invalid documents data from MVCR
        /// </summary>
        public static void DownloadMVCRData()
        {
            WebClient mvcrClient = new WebClient();
            mvcrClient.Headers.Add("Accept: text/html, application/xhtml+xml, */*");
            mvcrClient.Headers.Add("User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36");
            mvcrClient.DownloadFile(new Uri("https://aplikace.mvcr.cz/neplatne-doklady/ViewFile.aspx?typ_dokladu=0"), "op_vse.zip");
            mvcrClient.DownloadFile(new Uri("https://aplikace.mvcr.cz/neplatne-doklady/ViewFile.aspx?typ_dokladu=1"), "ops_vse.zip");
            mvcrClient.DownloadFile(new Uri("https://aplikace.mvcr.cz/neplatne-doklady/ViewFile.aspx?typ_dokladu=2"), "cd_vse.zip");
        }

        /// <summary>
        /// Extract .csv files from .zip files then deletes old .zip files
        /// </summary>
        /// <param name="zipPath"></param>
        public static void PrepareForReading(string zipPath)
        {
            using (ZipArchive zip = ZipFile.Open(zipPath, ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    if (entry.Name == "op_vse.csv")
                        entry.ExtractToFile("op_vse.csv");

                    if (entry.Name == "ops_vse.csv")
                        entry.ExtractToFile("ops_vse.csv");

                    if (entry.Name == "cd_vse.csv")
                        entry.ExtractToFile("cd_vse.csv");
                }
            }

            File.Delete(zipPath);
        }

        /// <summary>
        /// Deletes old files
        /// </summary>
        public static void DeleteOldCsvFiles()
        {
            File.Delete("op_vse.csv");
            File.Delete("ops_vse.csv");
            File.Delete("cd_vse.csv");
        }

        static int Main(string[] args)
        {
            if (args.Length < 2)
                return (int)ExitCode.Failure;

            string username = args[0];
            string pswd = args[1];
            DownloadMVCRData();
            if (File.Exists("op_vse.csv") || File.Exists("ops_vse.csv") || File.Exists("cd_vse.csv"))
                DeleteOldCsvFiles();
            PrepareForReading("op_vse.zip");
            PrepareForReading("ops_vse.zip");
            PrepareForReading("cd_vse.zip");
            Run(username, pswd);
            DeleteOldCsvFiles();

            return (int)ExitCode.Success;
        }
    }
}

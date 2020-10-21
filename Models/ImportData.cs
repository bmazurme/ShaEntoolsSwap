using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Entools.Model;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Entools.Models
{
    class ImportData
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API .NET EnTools";

        /// <summary>
        /// 
        /// </summary>
        /// <param name="revit"></param>
        public void ImportDataHydro()
        {
            ExternalCommandData revit = Transfer.revit;
            UIApplication uiApp = revit.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            string path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\credentials.json";
            UserCredential credential;

            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,
                Scopes, "user", CancellationToken.None, new FileDataStore(credPath + @"\entools", true)).Result;
            }
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            String spreadsheetId = Properties.Settings.Default["Sheet"].ToString();
            String range = "IMPORT!E2:I";
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            ValueRange response = request.Execute();

            IList<IList<Object>> values = response.Values;

            if (values != null && values.Count > 0)
            {
                try
                {
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("Change size");
                        foreach (var row in values)
                        {
                            int i = Convert.ToInt32(row[0].ToString().Substring(0, row[0].ToString().IndexOf("|")));
                            ElementId elementId = new ElementId(i);
                            Element element = doc.GetElement(elementId);

                            string flow = row[4].ToString().Substring(row[4].ToString().IndexOf("|")
                                                + 1, row[4].ToString().Length - row[4].ToString().IndexOf("|") - 1);
                            string velocity = row[0].ToString().Substring(row[0].ToString().IndexOf("|")
                                                + 1, row[0].ToString().Length - row[0].ToString().IndexOf("|") - 1);
                            string pressure = row[2].ToString().Substring(row[2].ToString().IndexOf("|")
                                                + 1, row[2].ToString().Length - row[2].ToString().IndexOf("|") - 1);

                            element.LookupParameter("entools_flow").Set(Convert.ToDouble(flow));
                            element.LookupParameter("entools_velocity").Set(velocity);
                            element.LookupParameter("entools_pressure").Set(pressure);
                        }
                        tx.Commit();
                    }
                }
                catch
                {
                    TaskDialog.Show("Error", "Check your data");
                    return;
                }
                TaskDialog.Show("Message", "Data imported!");
            }
            else
            {
                TaskDialog.Show("Message", "No data found.");
                return;
            }
        }
    }
}

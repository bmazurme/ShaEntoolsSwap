using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Entools.Model;
using Entools.Model.LibTools;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;

namespace Entools.Models
{
    class ClassHydroDataGS
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API .NET EnTools";

        public void ExportDataHydro()
        {
            ExternalCommandData revit = Transfer.revit;
            UIApplication uiApp = revit.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;

            List<BuiltInCategory> listCategories = new List<BuiltInCategory>()
            {
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_FlexPipeCurves
            };

            LibCategories libCategories = new LibCategories();
            List<Element> listElements = libCategories.AllElnewe(doc, listCategories).ToList();

            if (listElements[0].LookupParameter("entools_number") == null)
            {
                TaskDialog.Show("Error", "Parameter entools_number missing.");

                return;
            }

            listElements = listElements.OrderBy(x => int.Parse(x.LookupParameter("entools_number").AsValueString())).ToList();
            
            const BuiltInParameter bipSyst = BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM; 
            const BuiltInParameter bipDiameterInner = BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM;
            const BuiltInParameter bipLength = BuiltInParameter.CURVE_ELEM_LENGTH;
            const BuiltInParameter bipSize = BuiltInParameter.RBS_CALCULATED_SIZE;
            const BuiltInParameter minSize = BuiltInParameter.RBS_PIPE_SIZE_MINIMUM;
            const BuiltInParameter maxSize = BuiltInParameter.RBS_PIPE_SIZE_MAXIMUM;

            var listId = new List<object>();
            var listName = new List<object>();
            var listType = new List<object>();
            var listDo = new List<object>();
            var listDi = new List<object>();
            var listLength = new List<object>();
            var listKvs = new List<object>();

            try
            {
                Reference selection = sel.PickObject(ObjectType.Element, "Select element of system");
                Element element = doc.GetElement(selection);

                string stype = element.get_Parameter(bipSyst).AsString();
                string i = string.Empty;
                string getType = string.Empty;
                double getLen;
                double dinner;
                string size = string.Empty;
                string minsize = string.Empty;
                string maxsize = string.Empty;
                string id = string.Empty;

                listElements = listElements.Where(x => x.get_Parameter(bipSyst).AsString() == stype).ToList();

                foreach (Element pipe in listElements)
                {
                    if (pipe.LookupParameter("entools_number") == null)
                    {
                        TaskDialog.Show("Error", "Parameter entools_number missing.");
                        break;
                    }
                    else
                    {
                        getType = pipe.get_Parameter(bipSyst).AsString();
                        i = pipe.LookupParameter("entools_number").AsValueString();
                        size = pipe.get_Parameter(bipSize).AsString();

                        Category category = pipe.Category;
                        BuiltInCategory builtInCategory = (BuiltInCategory)category.Id.IntegerValue;

                        if (builtInCategory == BuiltInCategory.OST_PipeCurves)
                        {
                            id = pipe.Id.ToString();
                            getLen = pipe.get_Parameter(bipLength).AsDouble();
                            dinner = pipe.get_Parameter(bipDiameterInner).AsDouble();
                            getLen = Math.Round(UnitUtils.ConvertFromInternalUnits(getLen, DisplayUnitType.DUT_MILLIMETERS), 0);
                            dinner = Math.Round(UnitUtils.ConvertFromInternalUnits(dinner, DisplayUnitType.DUT_MILLIMETERS), 1);

                            listName.Add(getType + ":" + i.ToString());
                            listId.Add(pipe.Id.ToString());
                            listType.Add("PIP");
                            listDo.Add(size.Replace('.', ','));
                            listDi.Add(dinner.ToString().Replace('.', ','));
                            listLength.Add(getLen.ToString().Replace('.', ','));
                            listKvs.Add(pipe.LookupParameter("entools_kvse").AsValueString());
                        }

                        if (builtInCategory == BuiltInCategory.OST_PipeFitting)
                        {
                            minsize = pipe.get_Parameter(minSize).AsValueString();
                            maxsize = pipe.get_Parameter(maxSize).AsValueString();
                            listName.Add(getType + ":" + i.ToString());
                            listId.Add(pipe.Id.ToString());
                            listType.Add("FIT");
                            listDo.Add(maxsize.Replace('.', ',') + "-" + minsize.Replace('.', ','));
                            listDi.Add(minsize.Replace('.', ','));
                            listLength.Add("-");
                            listKvs.Add(pipe.LookupParameter("entools_kvse").AsValueString());
                        }

                        if (builtInCategory == BuiltInCategory.OST_PipeAccessory)
                        {
                            listName.Add(getType + ":" + i.ToString());
                            listId.Add(pipe.Id.ToString());
                            listType.Add("ACC");
                            listDo.Add(size);
                            listDi.Add("-");
                            listLength.Add("-");
                            listKvs.Add(pipe.LookupParameter("entools_kvse").AsValueString());
                        }

                        if (builtInCategory == BuiltInCategory.OST_FlexPipeCurves)
                        {
                            getLen = pipe.get_Parameter(bipLength).AsDouble();
                            dinner = pipe.get_Parameter(bipDiameterInner).AsDouble();
                            getLen = Math.Round(UnitUtils.ConvertFromInternalUnits(getLen, DisplayUnitType.DUT_MILLIMETERS), 0);
                            dinner = Math.Round(UnitUtils.ConvertFromInternalUnits(dinner, DisplayUnitType.DUT_MILLIMETERS), 1);

                            listName.Add(getType + ":" + i.ToString());
                            listId.Add(pipe.Id.ToString());
                            listType.Add("FLE");
                            listDo.Add(size.Replace('.', ','));
                            listDi.Add(dinner.ToString());
                            listLength.Add(getLen.ToString());
                            listKvs.Add("-");
                        }
                    }
                }

                ExportGD(listName, listType, listId, listDo, listDi, listLength, listKvs);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace);
                TaskDialog.Show("Report", "Canceled");
            }
        }


        private void ExportGD(List<object> namelist, List<object> typelist, List<object> oblist,
        List<object> doutlist, List<object> dinlist, List<object> lenlist, List<object> listkvs)
        {
            string path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\credentials.json";
            UserCredential credential;

            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,
                Scopes, "user_e", CancellationToken.None, new FileDataStore(credPath + @"\entools", true)).Result;
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            String spreadsheetId = Properties.Settings.Default["Sheet"].ToString();
            String range = "hydro!A2";

            ValueRange valueRange = new ValueRange
            {
                MajorDimension = "COLUMNS",
                Values = new List<IList<object>> { namelist, typelist, oblist, doutlist, dinlist, lenlist, listkvs }
            };

            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

            try
            {
                UpdateValuesResponse result = update.Execute();

                string url = Properties.Settings.Default["LinkUrl"].ToString();

                if (!string.IsNullOrEmpty(url))
                {
                    TaskDialog.Show("Message", "Data exported!");
                    System.Diagnostics.Process.Start(url);
                }
            }
            catch (Exception ex)
            {
                string error403 = "[403]";
                string error404 = "[404]";
                string error = ex.Message;
                
                if (error.IndexOf(error404) >= 0) MessageBox.Show("Please check the URL of the page.");

                if (error.IndexOf(error403) >= 0) MessageBox.Show("Please check the URL of the page and access rights.");
            }   
        }
    }
}
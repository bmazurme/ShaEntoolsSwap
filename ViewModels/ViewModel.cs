using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Entools.Model;
using Entools.Models;
using Entools.Plumbing.View;
using Entools.Views;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Input;

namespace Entools.ViewModels
{
    class ViewModel : INotifyPropertyChanged
    {
        #region FIELDS
        private bool _isEnableButton;
        public bool IsEnableButton
        {
            get
            {
                return _isEnableButton;
            }
            set
            {
                _isEnableButton = value;
                OnPropertyChanged("IsEnableButton");
            }
        }


        private string _addParStatus;
        public string AddParStatus
        {
            get { return _addParStatus; }
            set
            {
                if (_addParStatus != value)
                {
                    _addParStatus = value;
                    OnPropertyChanged("AddParStatus");
                }
            }
        }
        #endregion

        #region COMMANDS
        private ICommand _buttonOpenLink;
        public ICommand ButtonOpenLink
        {
            get
            {
                return _buttonOpenLink ?? (_buttonOpenLink = new CommandHandler(() => ButtonOpenLinkCommand(), () => CanExecute));
            }
        }

        private ICommand _buttonImport;
        public ICommand ButtonImport
        {
            get
            {
                return _buttonImport ?? (_buttonImport = new CommandHandler(() => ButtonImportCommand(), () => CanExecute));
            }
        }

        private ICommand _buttonExport;
        public ICommand ButtonExport
        {
            get
            {
                return _buttonExport ?? (_buttonExport = new CommandHandler(() => ButtonExportCommand(), () => CanExecute));
            }
        }

        private ICommand _buttonWindowLink;
        public ICommand ButtonWindowLink
        {
            get
            {
                return _buttonWindowLink ?? (_buttonWindowLink = new CommandHandler(() => ButtonWindowLinkCommand(), () => CanExecute));
            }
        }

        private ICommand AddParametersCommand;
        public ICommand ButtonAddEnPar
        {
            get
            {
                return AddParametersCommand ?? (AddParametersCommand = new CommandHandler(() => MyActionAddParameters(), () => CanExecute));
            }
        }
        #endregion

        #region ACTIONS
        public void ButtonOpenLinkCommand() 
        {
            string url = Properties.Settings.Default["LinkUrl"].ToString();

            if (!string.IsNullOrEmpty(url)) System.Diagnostics.Process.Start(url);
            else MessageBox.Show("URL empty");
        }

        public void ButtonImportCommand()
        {
            CloseAction();
            ImportData importData = new ImportData();
            importData.ImportDataHydro();
        }

        public void ButtonExportCommand()
        {
            string url = Properties.Settings.Default["LinkUrl"].ToString();

            if (!string.IsNullOrEmpty(url))
            {
                CloseAction();
                ClassHydroDataGS classHydroDataGS = new ClassHydroDataGS();
                classHydroDataGS.ExportDataHydro();
            }
        }

        public void ButtonWindowLinkCommand()
        {
            WindowLink windowLink = new WindowLink();
            windowLink.ShowDialog();
        }
        #endregion

        public void MyActionAddParameters()
        {
            #region INTERFACE

            CultureInfo ui = Thread.CurrentThread.CurrentUICulture;

            if (ui.Name == "ru-RU") System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("ru-RU");
            else System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

            #endregion

            ExternalCommandData revit = Transfer.revit;
            UIApplication uiapp = revit.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            ClassMillEnToolsParameters classMillEnToolsParameters = new ClassMillEnToolsParameters();
            string mes = classMillEnToolsParameters.CreateSharedParameter(uiapp, doc);
            if (!string.IsNullOrEmpty(mes)) AddParStatus = "Entools Swap - Parameters downloaded";
            else AddParStatus = "Failed";
        }

        public ViewModel()
        {
            _addParStatus = "Entools Swap";
            OnPropertyChanged("AddParStatus");
        }

        public Action CloseAction { get; set; }

        public bool CanExecute
        {
            get
            {
                // check if executing is allowed, i.e., validate, check if a process is running, etc. 
                //return true / false;
                return true;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }


    public class CommandHandler : ICommand
    {
        private Action _action;
        private Func<bool> _canExecute;

        /// <summary>
        /// Creates instance of the command handler
        /// </summary>
        /// <param name="action">Action to be executed by the command</param>
        /// <param name="canExecute">A bolean property to containing current permissions to execute the command</param>
        public CommandHandler(Action action, Func<bool> canExecute)
        {
            _action = action;
            _canExecute = canExecute;
        }


        /// <summary>
        /// Wires CanExecuteChanged event 
        /// </summary>
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }


        /// <summary>
        /// Forcess checking if execute is allowed
        /// </summary>
        /// <param name="parameter"></param>
        /// <returns></returns>
        public bool CanExecute(object parameter)
        {
            return _canExecute.Invoke();
        }


        public void Execute(object parameter)
        {
            _action();
        }
    }


    class ClassMillEnToolsParameters
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET EnTools";

        public static void Mainf(List<string> values)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Personal)
                                        + "/entools/EnTools_SharedParameters.txt";

            if (!File.Exists(path))
            {
                using (StreamWriter streamWriter = File.CreateText(path))
                {
                    foreach (var t in values) streamWriter.WriteLine(t);
                }
            }

            using (StreamReader streamReader = File.OpenText(path))
            {
                string s = string.Empty;
                while ((s = streamReader.ReadLine()) != null)
                {
                    Console.WriteLine(s);
                }
            }
        }


        private CategorySet Cat(Document doc, List<BuiltInCategory> listBuiltInCategory, CategorySet categories)
        {
            foreach (BuiltInCategory builtInCategory in listBuiltInCategory)
            {
                Category category = Category.GetCategory(doc, builtInCategory);
                categories.Insert(category);
            }

            return categories;
        }


        private void Insert(Document doc, UIApplication uiapp, Definition visibleParamDef, BindingMap bindingMap)
        {
            CategorySet categorySet = uiapp.Application.Create.NewCategorySet();

            List<BuiltInCategory> listBuiltInCategory = new List<BuiltInCategory>()
            {
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PipeInsulations,
                BuiltInCategory.OST_FlexPipeCurves,
                BuiltInCategory.OST_DetailComponents
            };

            categorySet = Cat(doc, listBuiltInCategory, categorySet);
            InstanceBinding typeBinding = uiapp.Application.Create.NewInstanceBinding(categorySet);
            bindingMap.Insert(visibleParamDef, typeBinding, BuiltInParameterGroup.INVALID);
        }


        private static bool CreateSharedParameterFile(Document doc, string paramFileDir, string paramFileName)
        {
            string paramFile = paramFileDir + paramFileName;

            if (File.Exists(paramFile))
            {
                doc.Application.SharedParametersFilename = paramFile;
                return true;
            }

            if (!Directory.Exists(paramFileDir)) Directory.CreateDirectory(paramFileDir);

            FileStream fileStream = File.Create(paramFile);
            fileStream.Close();
            doc.Application.SharedParametersFilename = paramFile;

            return true;
        }


        private void AddFileSharedParameter()
        {
            string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\credentials.json";
            UserCredential credential;

            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + @"\entools\";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes,
                                                    "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Id Google Sheets with parameters
            String spreadsheetId = "1w8tSwAxXk5UV7IFe6D7p5Jp9e2KB_ymuaLrp8uf-v-I";
            String range = "parameters!A1:I";
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;

            if (values != null && values.Count > 0)
            {
                List<string> text = new List<string>();

                foreach (var row in values)
                {
                    string line = string.Empty;

                    foreach (var e in row) line = line + e.ToString() + "	";
                    
                    text.Add(line);
                }

                Mainf(text);
            }
        }


        public string CreateSharedParameter(UIApplication uiApp, Document doc)
        {
            string message = string.Empty;
            AddFileSharedParameter();
            string definitionGroup = "EnTools";

            if (!CreateSharedParameterFile(doc, System.Environment
                .GetFolderPath(System.Environment.SpecialFolder.Personal)
                + @"\entools\", "EnTools_SharedParameters.txt"))
            {
                ErrorWindow Err = new ErrorWindow("!createSharedParameterFile", "Error");
                Err.ShowDialog();
            }

            List<string> names = new List<string>()
            {
                "_entools1", "_entools2", "_entools3", "_entools4", "_entools5",
                "_entools6", "_entools7", "_entools8", "_entools9", "_entools10",
                "entools_number", "entools_flow", "entools_velocity", "entools_pressure",
                "entools_kvse", "_entools_TD"
            };

            DefinitionFile parafile = doc.Application.OpenSharedParameterFile();
            CategorySet categories = doc.Application.Create.NewCategorySet();

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Copy Parameters");

                InstanceBinding binding = doc.Application.Create.NewInstanceBinding(categories);
                DefinitionGroup apiGroup = parafile.Groups.get_Item(definitionGroup);

                if (apiGroup == null) apiGroup = parafile.Groups.Create(definitionGroup);

                BindingMap bindingMap = doc.ParameterBindings;

                foreach (string name in names)
                {
                    Definition sharedParameterDefinition = apiGroup.Definitions.get_Item(name);
                    if (sharedParameterDefinition == null)
                    {
                        ExternalDefinitionCreationOptions option
                            = new ExternalDefinitionCreationOptions(name, ParameterType.Text);

                        Definition visibleParamDef = apiGroup.Definitions.Create(option);
                        Insert(doc, uiApp, visibleParamDef, bindingMap);
                    }
                    else
                    {
                        Definition visibleParamDef = sharedParameterDefinition;
                        Insert(doc, uiApp, visibleParamDef, bindingMap);
                    }
                    message = message + "Add parameter: " + name + "\n";
                }

                tx.Commit();
            }

            return message;
        }
    }
}
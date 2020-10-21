using System;
using System.IO;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace Entools.Model
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]

    public class Scheme : IExternalApplication
    {
        static string AddInPath = typeof(Entools).Assembly.Location;
        static string ButtonIconsFolder = Path.GetDirectoryName(AddInPath);

        public Autodesk.Revit.UI.Result OnStartup(UIControlledApplication application)
        {
            try
            {
                RibbonPanel panel = application.CreateRibbonPanel("EnTools");
                PushButtonData list = new PushButtonData("Swap", "Swap", AddInPath, "Entools.Model.Entools")
                {
                    ToolTip = "Swap"
                };
                string path = Path.GetDirectoryName( System.Reflection.Assembly.GetExecutingAssembly().Location);

                ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.ChmFile, path + @"\Help.htm"); // hard coding for simplicity. 
                list.SetContextualHelp(contextHelp);
                PushButton billButton = panel.AddItem(list) as PushButton;
                billButton.LargeImage = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "icons\\iconlarge.png"), UriKind.Absolute));
                billButton.Image = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "icons\\iconsmall.png"), UriKind.Absolute));

                return Autodesk.Revit.UI.Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("EnToolsLt Sample", ex.ToString());
                return Autodesk.Revit.UI.Result.Failed;
            }
        }

        public Autodesk.Revit.UI.Result OnShutdown(UIControlledApplication application)
        {
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
}
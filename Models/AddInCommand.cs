using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Entools.Views;

namespace Entools.Model
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]

    public class Entools : IExternalCommand
    {
        #region IExternalCommand Members Implementation
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit, ref string message, ElementSet elements)
        {
            try
            {
                // Open window
                Transfer.revit = revit;
                WindowScheme windowScheme = new WindowScheme();
                windowScheme.ShowDialog();

                return Autodesk.Revit.UI.Result.Succeeded;
            }
            catch (Exception ex)
            {
                // Error message
                message = ex.ToString();
                return Autodesk.Revit.UI.Result.Failed;
            }
        }
        #endregion IExternalCommand Members Implementation
    }
}
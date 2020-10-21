using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace Entools.Model
{
    public class SystemElSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            Category category = element.Category;
            BuiltInCategory enumCategory = (BuiltInCategory)category.Id.IntegerValue;

            if (enumCategory == BuiltInCategory.OST_PipeCurves
                || enumCategory == BuiltInCategory.OST_PipeFitting
                || enumCategory == BuiltInCategory.OST_PipeAccessory
                || enumCategory == BuiltInCategory.OST_PlumbingFixtures
                || enumCategory == BuiltInCategory.OST_MechanicalEquipment
                || enumCategory == BuiltInCategory.OST_FlexPipeCurves)
                return true;

            return false;
        }
        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }


    public static class Transfer
    {
        public static ExternalCommandData revit = null;
        public static ProgressBar progressBar = null;
    }
}
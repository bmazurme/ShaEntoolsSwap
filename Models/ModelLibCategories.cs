using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Entools.Model.LibTools
{
    class LibCategories
    {
        /// <summary>
        /// Get list of elements
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="builtInCategory"></param>
        /// <returns></returns>
        public IList<Element> AllElnew(Document doc, BuiltInCategory builtInCategory)
        {
            ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(builtInCategory);
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            IList<Element> listElements = collector.WherePasses(elementCategoryFilter)
                                            .WhereElementIsNotElementType().ToElements();
            return listElements;
        }


        public IList<Element> AllElnewe(Document doc, List<BuiltInCategory> listBuiltInCategories)
        {
            List<Element> listElements = new List<Element>();

            foreach (BuiltInCategory category in listBuiltInCategories)
            {
                ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(category);
                FilteredElementCollector filteredElementCollector = new FilteredElementCollector(doc);
                IList<Element> listDelta = filteredElementCollector.WherePasses(elementCategoryFilter)
                                                        .WhereElementIsNotElementType().ToElements();
                listElements.AddRange(listDelta);
            }
            return listElements;
        }

        /// <summary>
        /// Get list elements from active view
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="listBuiltInCategories"></param>
        /// <returns></returns>
        public IList<Element> AllElneweActiveView(Document doc, List<BuiltInCategory> listBuiltInCategories)
        {
            List<Element> listElements = new List<Element>();

            foreach (BuiltInCategory category in listBuiltInCategories)
            {
                ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(category);
                FilteredElementCollector filteredElementCollector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                IList<Element> list = filteredElementCollector.WherePasses(elementCategoryFilter)
                                                        .WhereElementIsNotElementType().ToElements();
                listElements.AddRange(list);
            }
            return listElements;
        }
    }
}
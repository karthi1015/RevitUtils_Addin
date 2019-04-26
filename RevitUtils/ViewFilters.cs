using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.Exceptions;

namespace RevitUtils
{
    static class ViewFilters
    {
        private static List<ElementId> GetUsedFilterIds(Document doc)
        {
            var views = new FilteredElementCollector(doc).OfClass(typeof(View)).ToElements().Cast<View>();
            List<ElementId> usedFilterIds = new List<ElementId>();

            foreach (var item in views)
            {
                ICollection<ElementId> viewFilterIds = null;
                try
                {
                    viewFilterIds = item.GetFilters();
                }
                catch (InvalidOperationException)
                {
                }

                if (viewFilterIds != null) usedFilterIds.AddRange(viewFilterIds);
            }

            return usedFilterIds;
        }

        public static ICollection<ElementId> GetUnUsedFilterIds(Document doc)
        {
            List<ElementId> usedFilterIds = GetUsedFilterIds(doc).ToList();

            var unusedFilterIds = usedFilterIds.Count > 0 ?
                new FilteredElementCollector(doc).OfClass(typeof(ParameterFilterElement)).Excluding(usedFilterIds).ToElementIds() :
                new FilteredElementCollector(doc).OfClass(typeof(ParameterFilterElement)).ToElementIds();

            return unusedFilterIds;
        }


        public static IList<Element> GetDocFilters(Document doc)
        {
            return new FilteredElementCollector(doc).OfClass(typeof(ParameterFilterElement)).ToElements();
        }
    }
}
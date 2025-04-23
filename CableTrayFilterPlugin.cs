using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace RevitFilterPlugin
{
    [Transaction(TransactionMode.Manual)]
    public class CreateCableTrayFilters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            View activeView = doc.ActiveView;

            if (!IsViewSupportsFilters(activeView))
            {
                TaskDialog.Show("错误", "当前视图类型不支持过滤器！");
                return Result.Failed;
            }

            var typeNames = CollectVisibleTypeNames(doc, activeView);
            if (typeNames.Count == 0) return Result.Succeeded;

            using (Transaction trans = new Transaction(doc, "创建桥架过滤器"))
            {
                try
                {
                    trans.Start();

                    var categories = new List<BuiltInCategory>
                    {
                        BuiltInCategory.OST_CableTray,
                        BuiltInCategory.OST_CableTrayFitting
                    };
                    var categoryIds = categories.Select(c => new ElementId(c)).ToList();

                    foreach (string typeName in typeNames)
                    {
                        string filterName = $"00-{typeName}";
                        if (FilterExists(doc, filterName)) continue;

                        // 创建参数过滤器
                        ParameterFilterElement paramFilter = ParameterFilterElement.Create(
                            doc,
                            filterName,
                            categoryIds);

                        // 创建过滤规则
                        FilterRule typeRule = ParameterFilterRuleFactory.CreateEqualsRule(
                            new ElementId(BuiltInParameter.SYMBOL_NAME_PARAM),
                            typeName,
                            true);

                        paramFilter.SetElementFilter(new ElementParameterFilter(typeRule));

                        // 应用过滤器到视图
                        if (!activeView.GetFilters().Contains(paramFilter.Id))
                        {
                            activeView.AddFilter(paramFilter.Id);
                        }
                    }

                    trans.Commit();
                    return Result.Succeeded;
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    TaskDialog.Show("错误", $"操作失败：{ex.Message}\n堆栈跟踪：{ex.StackTrace}");
                    return Result.Failed;
                }
            }
        }

        private HashSet<string> CollectVisibleTypeNames(Document doc, View view)
        {
            var typeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 创建多类别过滤器
            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting
            };
            var filter = new ElementMulticategoryFilter(categories);

            // 收集可见元素
            var elements = new FilteredElementCollector(doc, view.Id)
                .WherePasses(filter)
                .WhereElementIsNotElementType();

            foreach (Element instance in elements)
            {
                ElementType type = doc.GetElement(instance.GetTypeId()) as ElementType;
                if (type != null)
                {
                    Parameter param = type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                    if (param != null && !string.IsNullOrEmpty(param.AsString()))
                    {
                        typeNames.Add(param.AsString());
                    }
                }
            }

            return typeNames;
        }

        private bool FilterExists(Document doc, string filterName)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement))
                .Cast<ParameterFilterElement>()
                .Any(f => f.Name.Equals(filterName, StringComparison.OrdinalIgnoreCase));
        }

        private bool IsViewSupportsFilters(View view)
        {
            return view is ViewPlan || view is ViewSection || view is ViewDrafting;
        }
    }
}
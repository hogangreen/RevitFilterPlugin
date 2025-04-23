using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitFilterPlugin
{
    [Transaction(TransactionMode.Manual)]
    public class DustAndPipeFilterPlugin : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                View view = doc.ActiveView;
                string prefix = "00-";

                CreateDuctFilters(doc, view, prefix);
                CreatePipeFilters(doc, view, prefix);

                TaskDialog.Show("Success", "过滤器创建成功！");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("错误", ex.ToString());
                return Result.Failed;
            }
        }

        private void CreateDuctFilters(Document doc, View view, string prefix)
        {
            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_DuctTerminal
            };

            CreateSystemFilters(
                doc,
                view,
                prefix,
                categories,
                BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM
            );
        }

        private void CreatePipeFilters(Document doc, View view, string prefix)
        {
            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory
            };

            CreateSystemFilters(
                doc,
                view,
                prefix,
                categories,
                BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM
            );
        }

        private void CreateSystemFilters(
            Document doc,
            View view,
            string prefix,
            List<BuiltInCategory> categories,
            BuiltInParameter paramId)
        {
            var categoryIds = categories.ConvertAll(c => new ElementId(c));
            var multicategoryFilter = new ElementMulticategoryFilter(categoryIds);

            // 获取所有系统类型ID（参考桥架过滤器的实现）
            var systemIds = new FilteredElementCollector(doc)
                .WherePasses(multicategoryFilter)
                .Select(e =>
                {
                    var param = e.get_Parameter(paramId);
                    return param?.AsElementId() ?? ElementId.InvalidElementId;
                })
                .Where(id => id != ElementId.InvalidElementId)
                .Distinct()
                .ToList();

            using (Transaction trans = new Transaction(doc, "创建系统过滤器"))
            {
                trans.Start();

                foreach (ElementId systemId in systemIds)
                {
                    string systemName = doc.GetElement(systemId)?.Name ?? "Unknown";
                    string filterName = $"{prefix}{systemName}";

                    if (FilterExists(doc, filterName)) continue;

                    // 创建过滤器（使用与桥架相同的创建方式）
                    ParameterFilterElement filter = ParameterFilterElement.Create(
                        doc,
                        filterName,
                        categoryIds);

                    // 修正参数类型问题：使用ElementId比较规则
                    FilterRule rule = new FilterElementIdRule(
                        new ParameterValueProvider(new ElementId(paramId)), // 参数提供器
                        new FilterNumericEquals(),                          // 数值比较器
                        systemId                                            // 目标系统ID
                    );

                    filter.SetElementFilter(new ElementParameterFilter(rule));

                    if (!view.GetFilters().Contains(filter.Id))
                    {
                        view.AddFilter(filter.Id);
                    }
                }
                trans.Commit();
            }
        }

        private bool FilterExists(Document doc, string filterName)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement))
                .Any(x => x.Name.Equals(filterName));
        }
    }

    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            try
            {
                RibbonPanel panel = app.CreateRibbonPanel("MEP过滤器");
                panel.AddItem(new PushButtonData(
                    "CreateFilters",
                    "创建过滤器",
                    System.Reflection.Assembly.GetExecutingAssembly().Location,
                    "RevitFilterPlugin.DustAndPipeFilterPlugin")
                );
                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }
    }
}
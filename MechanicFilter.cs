using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RevitFilterPlugin
{
    [Transaction(TransactionMode.Manual)]
    public class MechanicFilter : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // 文档有效性验证
            if (doc == null || doc.IsLinked)
            {
                TaskDialog.Show("文档错误", "当前文档不可用或为链接模型");
                return Result.Cancelled;
            }

            View activeView = doc.ActiveView;
            if (activeView == null || activeView.IsTemplate)
            {
                TaskDialog.Show("视图错误", "当前视图不可用或为视图模板");
                return Result.Cancelled;
            }

            var targetCategories = new List<BuiltInCategory> {
                BuiltInCategory.OST_GenericModel,
                BuiltInCategory.OST_MechanicalEquipment
            };

            // 检测有效类别
            var existingCategories = targetCategories
                .Where(cat => new FilteredElementCollector(doc, activeView.Id)
                    .OfCategoryId(new ElementId(cat))
                    .Any(e => e?.Category != null))
                .ToList();

            if (existingCategories.Count == 0)
            {
                TaskDialog.Show("类别检测", "未找到目标类别元素");
                return Result.Cancelled;
            }

            using (Transaction trans = new Transaction(doc, "机械过滤器管理"))
            {
                trans.Start();

                try
                {
                    // 收集族数据
                    var familyData = CollectFamilyData(doc, activeView, existingCategories);
                    if (familyData.Count == 0)
                    {
                        TaskDialog.Show("数据错误", "未找到有效族数据");
                        trans.RollBack();
                        return Result.Cancelled;
                    }

                    int newFilters = 0;
                    int appliedFilters = 0;
                    ElementId paramId = new ElementId(BuiltInParameter.ALL_MODEL_FAMILY_NAME);

                    foreach (var data in familyData)
                    {
                        string filterName = $"M-{data.FamilyName}";

                        // 获取或创建过滤器
                        ParameterFilterElement filterElement = GetOrCreateFilter(
                            doc,
                            filterName,
                            data,
                            paramId,
                            ref newFilters);

                        // 应用过滤器到视图
                        if (ApplyFilterToView(activeView, filterElement))
                        {
                            appliedFilters++;
                        }
                    }

                    trans.Commit();

                    ShowResultDialog(doc, newFilters, appliedFilters);
                    return (newFilters + appliedFilters) > 0 ? Result.Succeeded : Result.Cancelled;
                }
                catch (Exception ex)
                {
                    HandleTransactionRollback(trans, ex);
                    return Result.Failed;
                }
            }
        }

        #region 核心逻辑方法
        private List<FamilyData> CollectFamilyData(Document doc, View view, List<BuiltInCategory> categories)
        {
            return new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .Where(e => e != null && e.Category != null)
                .Where(e => categories.Contains((BuiltInCategory)e.Category.Id.IntegerValue))
                .Select(e => new {
                    Element = e,
                    CategoryId = e.Category.Id
                })
                .Select(x => new FamilyData
                {
                    FamilyName = GetFamilyName(x.Element),  // 修正属性名
                    Categories = new List<ElementId> { x.CategoryId }  // 修正属性名
                })
                .Where(x => !string.IsNullOrEmpty(x.FamilyName))
                .GroupBy(x => x.FamilyName)
                .Select(g => new FamilyData
                {
                    FamilyName = g.Key,
                    Categories = g.SelectMany(x => x.Categories).Distinct().ToList()
                })
                .OrderBy(x => x.FamilyName)
                .ToList();
        }

        private ParameterFilterElement GetOrCreateFilter(
            Document doc,
            string filterName,
            FamilyData data,
            ElementId paramId,
            ref int newFilters)
        {
            // 尝试获取现有过滤器
            ParameterFilterElement filter = new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement))
                .Where(x => x.Name.Equals(filterName, StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault() as ParameterFilterElement;

            // 创建新过滤器
            if (filter == null)
            {
                FilterRule rule = ParameterFilterRuleFactory.CreateEqualsRule(
                    paramId,
                    data.FamilyName,  // 使用正确属性名
                    caseSensitive: false);

                filter = ParameterFilterElement.Create(
                    doc,
                    filterName,
                    data.Categories,  // 使用正确属性名
                    new ElementParameterFilter(rule));

                newFilters++;
            }

            return filter;
        }

        private bool ApplyFilterToView(View view, ParameterFilterElement filter)
        {
            if (view.GetFilters().Contains(filter.Id))
                return false;

            view.AddFilter(filter.Id);
            view.SetFilterVisibility(filter.Id, true);
            return true;
        }
        #endregion

        #region 辅助方法
        private string GetFamilyName(Element element)
        {
            try
            {
                if (element is FamilyInstance fi && fi.Symbol != null)
                    return fi.Symbol.FamilyName;

                Parameter p = element.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME);
                return p?.AsString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void ShowResultDialog(Document doc, int newFilters, int appliedFilters)
        {
            string viewName = doc.ActiveView?.Name ?? "未知视图";

            var dialog = new TaskDialog("过滤器管理报告")
            {
                TitleAutoPrefix = false,
                MainInstruction = $"操作结果 - {viewName}",
                MainContent = $"• 新建过滤器: {newFilters} 个\n" +
                              $"• 应用过滤器: {appliedFilters} 个\n" +
                              $"• 命名规则: M-族名称",
                FooterText = "提示：已存在的过滤器会被重新应用到当前视图",
                CommonButtons = TaskDialogCommonButtons.Ok
            };

            dialog.Show();
        }

        private void HandleTransactionRollback(Transaction trans, Exception ex)
        {
            if (trans.HasStarted() && !trans.HasEnded())
                trans.RollBack();

            TaskDialog.Show("系统错误",
                $"操作异常：{ex.Message}\n" +
                $"异常类型：{ex.GetType().Name}\n" +
                (ex.InnerException != null ? $"内部异常：{ex.InnerException.Message}" : ""));
        }
        #endregion

        #region 数据载体类
        private class FamilyData
        {
            public string FamilyName { get; set; }     // 正确属性名
            public List<ElementId> Categories { get; set; } // 正确属性名
        }
        #endregion
    }
}
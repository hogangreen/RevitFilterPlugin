using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;

namespace RevitFilterPlugin
{
    [Transaction(TransactionMode.Manual)]
    public class DeletePrefixFiltersCommand : IExternalCommand
    {
        private const string TargetPrefix = "00-"; // 定义要删除的过滤器前缀

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // 获取所有目标过滤器
                IList<ParameterFilterElement> targetFilters = GetTargetFilters(doc);

                if (targetFilters.Count == 0)
                {
                    TaskDialog.Show("提示", $"没有找到以'{TargetPrefix}'开头的过滤器");
                    return Result.Cancelled;
                }

                // 添加确认对话框
                TaskDialogResult confirmResult = TaskDialog.Show("确认删除",
                    $"即将删除 {targetFilters.Count} 个以'{TargetPrefix}'开头的过滤器，是否继续？",
                    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                if (confirmResult != TaskDialogResult.Yes)
                {
                    return Result.Cancelled;
                }

                using (Transaction trans = new Transaction(doc, "删除指定过滤器"))
                {
                    trans.Start();

                    try
                    {
                        int deletedCount = 0;
                        foreach (ParameterFilterElement filter in targetFilters)
                        {
                            if (IsValidForDeletion(doc, filter))
                            {
                                doc.Delete(filter.Id);
                                deletedCount++;
                            }
                        }

                        trans.Commit();
                        TaskDialog.Show("完成", $"已成功删除 {deletedCount} 个过滤器");
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        TaskDialog.Show("删除失败", ex.Message);
                        return Result.Failed;
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"操作失败：{ex.Message}";
                return Result.Failed;
            }
        }

        // 获取目标过滤器（兼容旧版）
        private IList<ParameterFilterElement> GetTargetFilters(Document doc)
        {
            List<ParameterFilterElement> filters = new List<ParameterFilterElement>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            foreach (Element element in collector.OfClass(typeof(ParameterFilterElement)))
            {
                if (element is ParameterFilterElement filter &&
                    filter.Name.StartsWith(TargetPrefix, StringComparison.Ordinal))
                {
                    filters.Add(filter);
                }
            }

            return filters;
        }

        // 安全删除验证
        private bool IsValidForDeletion(Document doc, ParameterFilterElement filter)
        {
            try
            {
                // 检查元素是否仍然存在
                return doc.GetElement(filter.Id) != null &&
                       filter.Name.StartsWith(TargetPrefix, StringComparison.Ordinal);
            }
            catch
            {
                return false;
            }
        }
    }
}
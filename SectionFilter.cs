using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RevitSectionFilter
{
    [Transaction(TransactionMode.Manual)]
    public class AdvancedSectionFilter : IExternalCommand
    {
        private const double ElevationTolerance = 0.001;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // 验证当前视图有效性 
                if (!(doc.ActiveView is ViewPlan currentView) || currentView.ViewType != ViewType.FloorPlan)
                {
                    TaskDialog.Show("操作终止", "请在楼层平面视图中运行此命令");
                    return Result.Cancelled;
                }

                // 获取当前视图关联标高 
                Level currentLevel = currentView.GenLevel ?? GetCurrentViewBottomLevel(doc, currentView);
                if (currentLevel == null)
                {
                    TaskDialog.Show("数据异常", "无法获取当前视图关联标高");
                    return Result.Failed;
                }

                // 收集所有有效剖面视图 
                List<ViewSection> sections = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSection))
                    .Cast<ViewSection>()
                    .Where(v => v.ViewType == ViewType.Section && !v.IsTemplate)
                    .ToList();

                // 执行隐藏操作
                var result = ProcessSections(doc, currentView, currentLevel, sections);

                // 刷新视图显示 
                uidoc.RefreshActiveView();

                // 显示操作报告
                ShowOperationResult(result, currentLevel.Name);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("系统错误", $"操作失败：{ex.Message}");
                return Result.Failed;
            }
        }

        private (int Total, int Hidden, List<string> Log) ProcessSections(
            Document doc,
            ViewPlan currentView,
            Level currentLevel,
            List<ViewSection> sections)
        {
            List<string> log = new List<string>();
            int hiddenCount = 0;

            // 获取所有标高数据 
            List<Level> allLevels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => l.Elevation)
                .ToList();

            using (Transaction trans = new Transaction(doc, "剖面标高过滤"))
            {
                trans.Start();

                foreach (ViewSection section in sections)
                {
                    try
                    {
                        // 获取剖面关联标高 
                        Level sectionLevel = GetSectionLevel(section, allLevels);

                        // 判断可见性 
                        bool shouldHide = sectionLevel?.Id != currentLevel.Id;

                        // 执行隐藏操作
                        if (shouldHide && TryHideSection(currentView, section))
                        {
                            hiddenCount++;
                            log.Add($"✅ 已隐藏：{section.Name} (ID:{section.Id.IntegerValue})");
                        }
                    }
                    catch (Exception ex)
                    {
                        log.Add($"❌ {section.Name} 处理失败：{ex.Message}");
                    }
                }

                trans.Commit();
            }

            return (sections.Count, hiddenCount, log);
        }

        private Level GetSectionLevel(ViewSection section, List<Level> levels)
        {
            // 精确计算剖面框中心点 
            BoundingBoxXYZ bbox = section.get_BoundingBox(section.Document.ActiveView);
            if (bbox != null)
            {
                XYZ midPoint = (bbox.Min + bbox.Max) * 0.5;
                return levels.OrderBy(l => Math.Abs(l.Elevation - midPoint.Z))
                    .FirstOrDefault(l => Math.Abs(l.Elevation - midPoint.Z) < ElevationTolerance);
            }

            // 备用方案：使用视图原点 
            return levels.OrderBy(l => Math.Abs(l.Elevation - section.Origin.Z))
                .FirstOrDefault(l => Math.Abs(l.Elevation - section.Origin.Z) < ElevationTolerance);
        }

        private bool TryHideSection(View view, ViewSection section)
        {
            // 检查元素是否已隐藏 
            if (section.IsHidden(view)) return false;

            // 检查隐藏权限 
            if (!section.CanBeHidden(view))
                throw new InvalidOperationException("元素受视图模板限制无法隐藏");

            view.HideElements(new List<ElementId> { section.Id });
            return true;
        }

        private Level GetCurrentViewBottomLevel(Document doc, ViewPlan view)
        {
            // 通过视图范围获取标高 
            try
            {
                PlanViewRange range = view.GetViewRange();
                ElementId levelId = range.GetLevelId(PlanViewPlane.BottomClipPlane);
                return doc.GetElement(levelId) as Level;
            }
            catch
            {
                return null;
            }
        }

        private void ShowOperationResult((int Total, int Hidden, List<string> Log) result, string levelName)
        {
            string summary = $"▸ 当前标高：{levelName}\n" +
                           $"▸ 处理剖面总数：{result.Total}\n" +
                           $"▸ 成功隐藏数量：{result.Hidden}\n" +
                           $"▸ 失败数量：{result.Log.Count - result.Hidden}";

            TaskDialog.Show("操作报告",
                $"{summary}\n\n详细日志：\n{string.Join("\n", result.Log.Take(10))}" +
                (result.Log.Count > 10 ? "\n......（更多记录请查看完整日志）" : ""));
        }
    }
}
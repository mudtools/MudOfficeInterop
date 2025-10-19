//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;

namespace ReportGenerationSystemSample
{
    /// <summary>
    /// 报表格式化器类
    /// </summary>
    public class ReportFormatter
    {
        /// <summary>
        /// 应用专业格式化到报表
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <returns>是否格式化成功</returns>
        public bool ApplyProfessionalFormatting(IWordDocument document)
        {
            try
            {
                // 设置页面布局
                using var pageSetup = document.Sections[1].PageSetup;
                pageSetup.PageSize = WdPaperSize.wdPaperA4;
                pageSetup.Orientation = WdOrientation.wdOrientPortrait;
                pageSetup.TopMargin = 1440;    // 2厘米
                pageSetup.BottomMargin = 1440;
                pageSetup.LeftMargin = 1800;   // 2.5厘米
                pageSetup.RightMargin = 1800;

                // 格式化标题
                FormatTitle(document);

                // 格式化表格
                FormatTables(document);

                // 格式化段落
                FormatParagraphs(document);

                Console.WriteLine("专业格式化已完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"格式化时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 格式化报表标题
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void FormatTitle(IWordDocument document)
        {
            try
            {
                // 查找并格式化标题
                using var find = document.Range().Find;
                find.ClearFormatting();
                find.Text = "XYZ公司月度销售报表";
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindContinue;

                if (find.Execute())
                {
                    find.ParentRange.Font.Name = "微软雅黑";
                    find.ParentRange.Font.Size = 20;
                    find.ParentRange.Font.Bold = true;
                    find.ParentRange.Font.Color = WdColor.wdColorDarkBlue;
                    find.ParentRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    find.ParentRange.ParagraphFormat.SpaceAfter = 24;
                }

                // 查找并格式化其他可能的标题
                find.Text = "XYZ公司月度财务报表";
                if (find.Execute())
                {
                    find.ParentRange.Font.Name = "微软雅黑";
                    find.ParentRange.Font.Size = 20;
                    find.ParentRange.Font.Bold = true;
                    find.ParentRange.Font.Color = WdColor.wdColorDarkBlue;
                    find.ParentRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    find.ParentRange.ParagraphFormat.SpaceAfter = 24;
                }

                find.Text = "XYZ公司项目进度报表";
                if (find.Execute())
                {
                    find.ParentRange.Font.Name = "微软雅黑";
                    find.ParentRange.Font.Size = 20;
                    find.ParentRange.Font.Bold = true;
                    find.ParentRange.Font.Color = WdColor.wdColorDarkBlue;
                    find.ParentRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    find.ParentRange.ParagraphFormat.SpaceAfter = 24;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"格式化标题时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 格式化报表中的表格
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void FormatTables(IWordDocument document)
        {
            try
            {
                // 格式化所有表格
                for (int i = 1; i <= document.Tables.Count; i++)
                {
                    using var table = document.Tables[i];

                    // 设置表格边框
                    table.Borders.Enable = true;
                    table.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    table.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                    table.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                    table.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;

                    // 设置表格对齐
                    table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

                    // 自动调整表格
                    table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"格式化表格时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 格式化报表中的段落
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void FormatParagraphs(IWordDocument document)
        {
            try
            {
                // 格式化正文段落
                foreach (var paragraph in document.Paragraphs)
                {
                    using var range = paragraph.Range;
                    if (range.Font.Size == 12 && range.Font.Name == "宋体")
                    {
                        range.ParagraphFormat.LineSpacing = 1.5f; // 1.5倍行距
                        range.ParagraphFormat.SpaceAfter = 12;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"格式化段落时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 应用公司品牌格式化
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="companyName">公司名称</param>
        /// <param name="primaryColor">主要颜色</param>
        /// <returns>是否格式化成功</returns>
        public bool ApplyBrandingFormatting(IWordDocument document, string companyName, WdColor primaryColor)
        {
            try
            {
                // 设置页眉中的公司名称
                using var headerRange = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Text = $"{companyName}报表";
                headerRange.Font.Name = "微软雅黑";
                headerRange.Font.Size = 14;
                headerRange.Font.Bold = true;
                headerRange.Font.Color = primaryColor;
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 设置标题格式
                using var titleRange = document.Range(0, 50); // 假设标题在前50个字符内
                titleRange.Font.Name = "微软雅黑";
                titleRange.Font.Size = 20;
                titleRange.Font.Bold = true;
                titleRange.Font.Color = primaryColor;
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                titleRange.ParagraphFormat.SpaceAfter = 24;

                Console.WriteLine($"品牌格式化已完成: {companyName}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用品牌格式化时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 应用现代风格格式化
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <returns>是否格式化成功</returns>
        public bool ApplyModernFormatting(IWordDocument document)
        {
            // 设置页面布局
            var pageSetup = document.Sections[1].PageSetup;
            try
            {
                pageSetup.PageSize = WdPaperSize.wdPaperA4;
                pageSetup.Orientation = WdOrientation.wdOrientPortrait;
                pageSetup.TopMargin = 1134;    // 1.5厘米
                pageSetup.BottomMargin = 1134;
                pageSetup.LeftMargin = 1701;   // 2.3厘米
                pageSetup.RightMargin = 1701;

                // 设置字体
                document.Range().Font.Name = "Segoe UI";
                document.Range().Font.Size = 11;

                // 格式化标题
                FormatModernTitle(document);

                // 格式化表格
                FormatModernTables(document);

                // 格式化段落
                FormatModernParagraphs(document);

                Console.WriteLine("现代风格格式化已完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"现代风格格式化时出错: {ex.Message}");
                return false;
            }
            finally
            {
                pageSetup?.Dispose();
            }
        }

        /// <summary>
        /// 格式化现代风格的标题
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void FormatModernTitle(IWordDocument document)
        {
            try
            {
                using var find = document.Range().Find;
                find.ClearFormatting();
                find.Text = "XYZ公司*";
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindContinue;
                find.MatchWildcards = true;

                while (find.Execute())
                {
                    find.ParentRange.Font.Name = "Segoe UI";
                    find.ParentRange.Font.Size = 24;
                    find.ParentRange.Font.Bold = true;
                    find.ParentRange.Font.Color = WdColor.wdColorDarkBlue;
                    find.ParentRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    find.ParentRange.ParagraphFormat.SpaceAfter = 20;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"格式化现代风格标题时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 格式化现代风格的表格
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void FormatModernTables(IWordDocument document)
        {
            try
            {
                for (int i = 1; i <= document.Tables.Count; i++)
                {
                    using var table = document.Tables[i];

                    // 设置无边框样式
                    table.Borders.Enable = false;

                    // 设置表头样式
                    if (table.Rows.Count > 0)
                    {
                        var headerRow = table.Rows[1];
                        headerRow.Range.Font.Bold = true;
                        headerRow.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        headerRow.Shading.BackgroundPatternColor = WdColor.wdColorLightBlue;

                        // 设置表头底边框
                        headerRow.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                        headerRow.Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth150pt;
                        headerRow.Borders[WdBorderType.wdBorderBottom].Color = WdColor.wdColorDarkBlue;
                    }

                    // 设置数据行样式
                    for (int j = 2; j <= table.Rows.Count; j++)
                    {
                        var row = table.Rows[j];
                        row.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        row.HeightRule = WdRowHeightRule.wdRowHeightAtLeast;
                        row.Height = 360; // 0.5厘米
                    }

                    // 自动调整表格
                    table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"格式化现代风格表格时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 格式化现代风格的段落
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void FormatModernParagraphs(IWordDocument document)
        {
            try
            {
                foreach (var paragraph in document.Paragraphs)
                {
                    var range = paragraph.Range;
                    range.ParagraphFormat.LineSpacing = 1.3f;
                    range.ParagraphFormat.SpaceAfter = 10;
                    range.ParagraphFormat.SpaceBefore = 6;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"格式化现代风格段落时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 应用简洁风格格式化
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <returns>是否格式化成功</returns>
        public bool ApplyMinimalistFormatting(IWordDocument document)
        {
            var pageSetup = document.Sections[1].PageSetup;
            try
            {
                // 设置页面布局
                pageSetup.PageSize = WdPaperSize.wdPaperA4;
                pageSetup.Orientation = WdOrientation.wdOrientPortrait;
                pageSetup.TopMargin = 1417;    // 2厘米
                pageSetup.BottomMargin = 1417;
                pageSetup.LeftMargin = 1417;
                pageSetup.RightMargin = 1417;

                // 设置字体
                document.Range().Font.Name = "Calibri";
                document.Range().Font.Size = 11;

                // 格式化标题
                FormatMinimalistTitle(document);

                // 格式化表格
                FormatMinimalistTables(document);

                // 格式化段落
                FormatMinimalistParagraphs(document);

                Console.WriteLine("简洁风格格式化已完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"简洁风格格式化时出错: {ex.Message}");
                return false;
            }
            finally
            {
                pageSetup?.Dispose();
            }
        }

        /// <summary>
        /// 格式化简洁风格的标题
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void FormatMinimalistTitle(IWordDocument document)
        {
            try
            {
                using var find = document.Range().Find;
                find.ClearFormatting();
                find.Text = "XYZ公司*";
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindContinue;
                find.MatchWildcards = true;

                while (find.Execute())
                {
                    find.ParentRange.Font.Name = "Calibri";
                    find.ParentRange.Font.Size = 18;
                    find.ParentRange.Font.Bold = true;
                    find.ParentRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    find.ParentRange.ParagraphFormat.SpaceAfter = 16;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"格式化简洁风格标题时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 格式化简洁风格的表格
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void FormatMinimalistTables(IWordDocument document)
        {
            try
            {
                for (int i = 1; i <= document.Tables.Count; i++)
                {
                    using var table = document.Tables[i];

                    // 设置简单边框
                    table.Borders.Enable = true;
                    table.Borders.LineStyle = WdLineStyle.wdLineStyleSingle;
                    table.Borders.LineWidth = WdLineWidth.wdLineWidth050pt;

                    // 设置表头样式
                    if (table.Rows.Count > 0)
                    {
                        using var headerRow = table.Rows[1];
                        headerRow.Range.Font.Bold = true;
                        headerRow.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        headerRow.Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    }

                    // 设置数据行样式
                    for (int j = 2; j <= table.Rows.Count; j++)
                    {
                        using var row = table.Rows[j];
                        row.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    }

                    // 自动调整表格
                    table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"格式化简洁风格表格时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 格式化简洁风格的段落
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void FormatMinimalistParagraphs(IWordDocument document)
        {
            try
            {
                foreach (var paragraph in document.Paragraphs)
                {
                    using var range = paragraph.Range;
                    range.ParagraphFormat.LineSpacing = 1.15f;
                    range.ParagraphFormat.SpaceAfter = 8;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"格式化简洁风格段落时出错: {ex.Message}");
            }
        }
    }
}
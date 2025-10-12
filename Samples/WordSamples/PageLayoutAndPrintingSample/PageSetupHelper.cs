using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PageLayoutAndPrintingSample
{
    /// <summary>
    /// 页面设置助手类
    /// </summary>
    public class PageSetupHelper
    {
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public PageSetupHelper(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 设置页面尺寸
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="paperSize">纸张尺寸</param>
        public void SetPaperSize(int sectionIndex, WdPaperSize paperSize)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    var section = _document.Sections[sectionIndex];
                    section.PageSetup.PageSize = paperSize;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置纸张尺寸时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 自定义设置页面尺寸
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="width">宽度（磅）</param>
        /// <param name="height">高度（磅）</param>
        public void SetCustomPageSize(int sectionIndex, float width, float height)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    var section = _document.Sections[sectionIndex];
                    section.PageSetup.PageWidth = width;
                    section.PageSetup.PageHeight = height;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置自定义页面尺寸时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 设置页面方向
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="orientation">页面方向</param>
        public void SetPageOrientation(int sectionIndex, WdOrientation orientation)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    var section = _document.Sections[sectionIndex];
                    section.PageSetup.Orientation = orientation;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置页面方向时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 设置页面边距
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="top">上边距（磅）</param>
        /// <param name="bottom">下边距（磅）</param>
        /// <param name="left">左边距（磅）</param>
        /// <param name="right">右边距（磅）</param>
        public void SetPageMargins(int sectionIndex, float top, float bottom, float left, float right)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    var section = _document.Sections[sectionIndex];
                    section.PageSetup.TopMargin = top;
                    section.PageSetup.BottomMargin = bottom;
                    section.PageSetup.LeftMargin = left;
                    section.PageSetup.RightMargin = right;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置页面边距时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 设置页眉页脚距离
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="headerDistance">页眉距离（磅）</param>
        /// <param name="footerDistance">页脚距离（磅）</param>
        public void SetHeaderFooterDistance(int sectionIndex, float headerDistance, float footerDistance)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    var section = _document.Sections[sectionIndex];
                    section.PageSetup.HeaderDistance = headerDistance;
                    section.PageSetup.FooterDistance = footerDistance;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置页眉页脚距离时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 设置页面垂直对齐方式
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="verticalAlignment">垂直对齐方式</param>
        public void SetVerticalAlignment(int sectionIndex, WdVerticalAlignment verticalAlignment)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    var section = _document.Sections[sectionIndex];
                    section.PageSetup.VerticalAlignment = verticalAlignment;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置页面垂直对齐方式时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 启用或禁用首页不同的页眉页脚
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="different">是否不同</param>
        public void SetDifferentFirstPageHeaderFooter(int sectionIndex, bool different)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    var section = _document.Sections[sectionIndex];
                    section.PageSetup.DifferentFirstPageHeaderFooter = different ? 1 : 0;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置首页不同页眉页脚时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 启用或禁用奇偶页不同的页眉页脚
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="different">是否不同</param>
        public void SetOddAndEvenPagesHeaderFooter(int sectionIndex, bool different)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    var section = _document.Sections[sectionIndex];
                    section.PageSetup.OddAndEvenPagesHeaderFooter = different ? 1 : 0;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置奇偶页不同页眉页脚时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 设置行号
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="active">是否启用行号</param>
        /// <param name="restartMode">重新开始模式</param>
        public void SetLineNumbering(int sectionIndex, bool active, WdNumberingRule restartMode)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    var section = _document.Sections[sectionIndex];
                    section.PageSetup.LineNumbering.Active = active ? 1 : 0;
                    section.PageSetup.LineNumbering.RestartMode = restartMode;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置行号时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取指定节的页面设置信息
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <returns>页面设置信息字符串</returns>
        public string GetPageSetupInfo(int sectionIndex)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    var section = _document.Sections[sectionIndex];
                    var pageSetup = section.PageSetup;

                    StringBuilder info = new StringBuilder();
                    info.AppendLine($"第 {sectionIndex} 节页面设置信息:");
                    info.AppendLine($"  纸张尺寸: {pageSetup.PageSize}");
                    info.AppendLine($"  页面宽度: {pageSetup.PageWidth} 磅");
                    info.AppendLine($"  页面高度: {pageSetup.PageHeight} 磅");
                    info.AppendLine($"  页面方向: {pageSetup.Orientation}");
                    info.AppendLine($"  上边距: {pageSetup.TopMargin} 磅");
                    info.AppendLine($"  下边距: {pageSetup.BottomMargin} 磅");
                    info.AppendLine($"  左边距: {pageSetup.LeftMargin} 磅");
                    info.AppendLine($"  右边距: {pageSetup.RightMargin} 磅");
                    info.AppendLine($"  页眉距离: {pageSetup.HeaderDistance} 磅");
                    info.AppendLine($"  页脚距离: {pageSetup.FooterDistance} 磅");
                    info.AppendLine($"  垂直对齐: {pageSetup.VerticalAlignment}");
                    info.AppendLine($"  首页不同: {pageSetup.DifferentFirstPageHeaderFooter == 1}");
                    info.AppendLine($"  奇偶页不同: {pageSetup.OddAndEvenPagesHeaderFooter == 1}");
                    info.AppendLine($"  行号启用: {pageSetup.LineNumbering.Active == 1}");

                    return info.ToString();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取页面设置信息时出错: {ex.Message}");
            }

            return $"无法获取第 {sectionIndex} 节的页面设置信息";
        }

        /// <summary>
        /// 复制页面设置到其他节
        /// </summary>
        /// <param name="sourceSectionIndex">源节索引</param>
        /// <param name="targetSectionIndices">目标节索引数组</param>
        public void CopyPageSetupToSections(int sourceSectionIndex, int[] targetSectionIndices)
        {
            try
            {
                if (sourceSectionIndex > 0 && sourceSectionIndex <= _document.Sections.Count)
                {
                    var sourcePageSetup = _document.Sections[sourceSectionIndex].PageSetup;

                    foreach (int targetIndex in targetSectionIndices)
                    {
                        if (targetIndex > 0 && targetIndex <= _document.Sections.Count && targetIndex != sourceSectionIndex)
                        {
                            var targetPageSetup = _document.Sections[targetIndex].PageSetup;
                            
                            // 复制所有页面设置属性
                            targetPageSetup.PageSize = sourcePageSetup.PageSize;
                            targetPageSetup.PageWidth = sourcePageSetup.PageWidth;
                            targetPageSetup.PageHeight = sourcePageSetup.PageHeight;
                            targetPageSetup.Orientation = sourcePageSetup.Orientation;
                            targetPageSetup.TopMargin = sourcePageSetup.TopMargin;
                            targetPageSetup.BottomMargin = sourcePageSetup.BottomMargin;
                            targetPageSetup.LeftMargin = sourcePageSetup.LeftMargin;
                            targetPageSetup.RightMargin = sourcePageSetup.RightMargin;
                            targetPageSetup.HeaderDistance = sourcePageSetup.HeaderDistance;
                            targetPageSetup.FooterDistance = sourcePageSetup.FooterDistance;
                            targetPageSetup.VerticalAlignment = sourcePageSetup.VerticalAlignment;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"复制页面设置时出错: {ex.Message}");
            }
        }
    }
}
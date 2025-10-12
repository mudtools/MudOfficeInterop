using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskPanesAndDialogsSample
{
    /// <summary>
    /// 对话框管理器类
    /// </summary>
    public class DialogManager
    {
        private readonly IWordApplication _application;
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="application">Word应用程序对象</param>
        /// <param name="document">Word文档对象</param>
        public DialogManager(IWordApplication application, IWordDocument document)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 显示字体对话框
        /// </summary>
        /// <returns>用户是否点击了确定</returns>
        public bool ShowFontDialog()
        {
            try
            {
                var fontDialog = _application.Dialogs[WdWordDialog.wdDialogFormatFont];
                int result = fontDialog.Show();

                if (result == 1) // 用户点击了确定
                {
                    Console.WriteLine("字体设置已应用");
                    return true;
                }
                else
                {
                    Console.WriteLine("用户取消了字体设置");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"显示字体对话框时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 显示段落对话框
        /// </summary>
        /// <returns>用户是否点击了确定</returns>
        public bool ShowParagraphDialog()
        {
            try
            {
                var paragraphDialog = _application.Dialogs[WdWordDialog.wdDialogFormatParagraph];
                int result = paragraphDialog.Show();

                if (result == 1) // 用户点击了确定
                {
                    Console.WriteLine("段落格式已应用");
                    return true;
                }
                else
                {
                    Console.WriteLine("用户取消了段落格式设置");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"显示段落对话框时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 显示页面设置对话框
        /// </summary>
        /// <returns>用户是否点击了确定</returns>
        public bool ShowPageSetupDialog()
        {
            try
            {
                var pageSetupDialog = _application.Dialogs[WdWordDialog.wdDialogFilePageSetup];
                int result = pageSetupDialog.Show();

                if (result == 1) // 用户点击了确定
                {
                    Console.WriteLine("页面设置已应用");
                    return true;
                }
                else
                {
                    Console.WriteLine("用户取消了页面设置");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"显示页面设置对话框时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 显示查找替换对话框
        /// </summary>
        /// <returns>用户是否点击了确定</returns>
        public bool ShowFindDialog()
        {
            try
            {
                var findDialog = _application.Dialogs[WdWordDialog.wdDialogEditFind];
                int result = findDialog.Show();

                if (result == 1) // 用户点击了确定
                {
                    Console.WriteLine("查找操作已完成");
                    return true;
                }
                else
                {
                    Console.WriteLine("用户取消了查找操作");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"显示查找对话框时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 显示自定义字体对话框并应用设置
        /// </summary>
        /// <returns>字体设置信息</returns>
        public FontSettings ShowCustomFontDialog()
        {
            var fontSettings = new FontSettings();

            try
            {
                var fontDialog = _application.Dialogs[WdWordDialog.wdDialogFormatFont];
                fontDialog.DefaultTab = WdFontDialogTab.wdFontTabFont;

                // 显示对话框并获取结果
                int result = fontDialog.Show();

                if (result == 1) // 用户点击了确定
                {
                    // 获取用户选择的设置
                    fontSettings.FontName = fontDialog.Font;
                    fontSettings.FontSize = fontDialog.Points;
                    fontSettings.IsBold = fontDialog.Bold != 0;
                    fontSettings.IsItalic = fontDialog.Italic != 0;

                    Console.WriteLine($"用户选择了字体: {fontSettings.FontName}, 大小: {fontSettings.FontSize}");
                    Console.WriteLine($"粗体: {fontSettings.IsBold}, 斜体: {fontSettings.IsItalic}");

                    // 应用到当前选择
                    var selection = _application.Selection;
                    if (selection != null)
                    {
                        selection.Font.Name = fontSettings.FontName;
                        selection.Font.Size = fontSettings.FontSize;
                        selection.Font.Bold = fontSettings.IsBold ? 1 : 0;
                        selection.Font.Italic = fontSettings.IsItalic ? 1 : 0;
                    }

                    fontSettings.IsApplied = true;
                }
                else
                {
                    Console.WriteLine("用户取消了字体设置");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"显示自定义字体对话框时出错: {ex.Message}");
                fontSettings.ErrorMessage = ex.Message;
            }

            return fontSettings;
        }

        /// <summary>
        /// 显示自定义页面设置对话框并应用设置
        /// </summary>
        /// <returns>页面设置信息</returns>
        public PageSettings ShowCustomPageSetupDialog()
        {
            var pageSettings = new PageSettings();

            try
            {
                var pageSetupDialog = _application.Dialogs[WdWordDialog.wdDialogFilePageSetup];

                // 显示对话框
                int result = pageSetupDialog.Show();

                if (result == 1) // 用户点击了确定
                {
                    // 应用页面设置到当前节
                    var section = _document.Sections[1];
                    var pageSetup = section.PageSetup;

                    pageSettings.PageWidth = pageSetup.PageWidth;
                    pageSettings.PageHeight = pageSetup.PageHeight;
                    pageSettings.TopMargin = pageSetup.TopMargin;
                    pageSettings.BottomMargin = pageSetup.BottomMargin;
                    pageSettings.LeftMargin = pageSetup.LeftMargin;
                    pageSettings.RightMargin = pageSetup.RightMargin;

                    Console.WriteLine("页面设置已更新");
                    Console.WriteLine($"页面宽度: {pageSettings.PageWidth}");
                    Console.WriteLine($"页面高度: {pageSettings.PageHeight}");
                    Console.WriteLine($"上边距: {pageSettings.TopMargin}");
                    Console.WriteLine($"下边距: {pageSettings.BottomMargin}");

                    pageSettings.IsApplied = true;
                }
                else
                {
                    Console.WriteLine("用户取消了页面设置");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"显示自定义页面设置对话框时出错: {ex.Message}");
                pageSettings.ErrorMessage = ex.Message;
            }

            return pageSettings;
        }

        /// <summary>
        /// 获取所有可用对话框信息
        /// </summary>
        /// <returns>对话框信息列表</returns>
        public List<DialogInfo> GetAllDialogsInfo()
        {
            var dialogInfos = new List<DialogInfo>();

            try
            {
                var dialogs = _application.Dialogs;
                int dialogCount = dialogs.Count;

                Console.WriteLine($"当前可用对话框数量: {dialogCount}");

                // 添加一些常用对话框信息
                dialogInfos.Add(new DialogInfo
                {
                    Id = (int)WdWordDialog.wdDialogFormatFont,
                    Name = "字体对话框",
                    Description = "用于设置字体格式"
                });

                dialogInfos.Add(new DialogInfo
                {
                    Id = (int)WdWordDialog.wdDialogFormatParagraph,
                    Name = "段落对话框",
                    Description = "用于设置段落格式"
                });

                dialogInfos.Add(new DialogInfo
                {
                    Id = (int)WdWordDialog.wdDialogFilePageSetup,
                    Name = "页面设置对话框",
                    Description = "用于设置页面布局"
                });

                dialogInfos.Add(new DialogInfo
                {
                    Id = (int)WdWordDialog.wdDialogEditFind,
                    Name = "查找对话框",
                    Description = "用于查找和替换文本"
                });

                foreach (var info in dialogInfos)
                {
                    Console.WriteLine($"对话框: {info.Name} (ID: {info.Id}) - {info.Description}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取对话框信息时出错: {ex.Message}");
            }

            return dialogInfos;
        }

        /// <summary>
        /// 生成对话框交互最佳实践指南
        /// </summary>
        /// <returns>最佳实践指南</returns>
        public string GenerateDialogBestPractices()
        {
            var guideBuilder = new StringBuilder();
            guideBuilder.AppendLine("=== 对话框交互最佳实践指南 ===");
            guideBuilder.AppendLine("");
            guideBuilder.AppendLine("1. 对话框设计原则:");
            guideBuilder.AppendLine("   - 使用标准对话框以保持一致性");
            guideBuilder.AppendLine("   - 提供明确的确定/取消按钮");
            guideBuilder.AppendLine("   - 验证用户输入");
            guideBuilder.AppendLine("   - 记住用户偏好设置");
            guideBuilder.AppendLine("");
            guideBuilder.AppendLine("2. 用户体验优化:");
            guideBuilder.AppendLine("   - 提供有意义的默认值");
            guideBuilder.AppendLine("   - 给出清晰的操作反馈");
            guideBuilder.AppendLine("   - 处理异常情况");
            guideBuilder.AppendLine("   - 支持撤销操作");
            guideBuilder.AppendLine("");
            guideBuilder.AppendLine("3. 错误处理:");
            guideBuilder.AppendLine("   - 捕获并处理对话框操作异常");
            guideBuilder.AppendLine("   - 提供友好的错误提示信息");
            guideBuilder.AppendLine("   - 记录错误日志便于调试");

            Console.WriteLine("已生成对话框交互最佳实践指南");
            return guideBuilder.ToString();
        }
    }

    /// <summary>
    /// 字体设置类
    /// </summary>
    public class FontSettings
    {
        /// <summary>
        /// 字体名称
        /// </summary>
        public string FontName { get; set; }

        /// <summary>
        /// 字体大小
        /// </summary>
        public int FontSize { get; set; }

        /// <summary>
        /// 是否粗体
        /// </summary>
        public bool IsBold { get; set; }

        /// <summary>
        /// 是否斜体
        /// </summary>
        public bool IsItalic { get; set; }

        /// <summary>
        /// 是否已应用
        /// </summary>
        public bool IsApplied { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 页面设置类
    /// </summary>
    public class PageSettings
    {
        /// <summary>
        /// 页面宽度
        /// </summary>
        public float PageWidth { get; set; }

        /// <summary>
        /// 页面高度
        /// </summary>
        public float PageHeight { get; set; }

        /// <summary>
        /// 上边距
        /// </summary>
        public float TopMargin { get; set; }

        /// <summary>
        /// 下边距
        /// </summary>
        public float BottomMargin { get; set; }

        /// <summary>
        /// 左边距
        /// </summary>
        public float LeftMargin { get; set; }

        /// <summary>
        /// 右边距
        /// </summary>
        public float RightMargin { get; set; }

        /// <summary>
        /// 是否已应用
        /// </summary>
        public bool IsApplied { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 对话框信息类
    /// </summary>
    public class DialogInfo
    {
        /// <summary>
        /// 对话框ID
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// 对话框名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 对话框描述
        /// </summary>
        public string Description { get; set; }
    }
}
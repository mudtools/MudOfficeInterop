//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace SelectionOperationsSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 选择区域操作示例");

            // 示例1: IWordSelection接口详解
            Console.WriteLine("\n=== 示例1: IWordSelection接口详解 ===");
            WordSelectionDemo();

            // 示例2: 选择区域的基本概念
            Console.WriteLine("\n=== 示例2: 选择区域的基本概念 ===");
            SelectionConceptsDemo();

            // 示例3: 选择区域类型和属性
            Console.WriteLine("\n=== 示例3: 选择区域类型和属性 ===");
            SelectionTypesDemo();

            // 示例4: 文本选择和操作
            Console.WriteLine("\n=== 示例4: 文本选择和操作 ===");
            TextSelectionDemo();

            // 示例5: 选择区域的扩展和收缩
            Console.WriteLine("\n=== 示例5: 选择区域的扩展和收缩 ===");
            SelectionExpansionDemo();

            // 示例6: 高级选择操作
            Console.WriteLine("\n=== 示例6: 高级选择操作 ===");
            AdvancedSelectionDemo();

            // 示例7: 选择区域格式化
            Console.WriteLine("\n=== 示例7: 选择区域格式化 ===");
            SelectionFormattingDemo();

            // 示例8: 使用辅助类的完整示例
            Console.WriteLine("\n=== 示例8: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// IWordSelection接口详解示例
        /// </summary>
        static void WordSelectionDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var selection = app.Selection;

                if (selection != null)
                {
                    // 获取选择区域文本
                    string selectedText = selection.Text;

                    // 获取选择区域类型
                    WdSelectionType selectionType = selection.Type;

                    // 检查选择区域是否处于活动状态
                    bool isActive = selection.Active;

                    Console.WriteLine($"选择区域文本: '{selectedText}'");
                    Console.WriteLine($"选择区域类型: {selectionType}");
                    Console.WriteLine($"选择区域是否活动: {isActive}");
                }

                Console.WriteLine("IWordSelection接口操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"IWordSelection接口操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 选择区域的基本概念示例
        /// </summary>
        static void SelectionConceptsDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var selection = app.Selection;

                if (selection != null)
                {
                    Console.WriteLine($"选择区域是否活动: {selection.Active}");
                    Console.WriteLine($"选择区域类型: {selection.Type}");
                    Console.WriteLine($"故事类型: {selection.StoryType}");
                    Console.WriteLine($"故事长度: {selection.StoryLength}");
                    Console.WriteLine($"是否在行尾: {selection.IPAtEndOfLine}");
                    Console.WriteLine($"是否在行末标记处: {selection.IsEndOfRowMark}");
                }

                Console.WriteLine("选择区域基本概念演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"选择区域基本概念演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 选择区域类型和属性示例
        /// </summary>
        static void SelectionTypesDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;
                var selection = app.Selection;

                // 添加一些文本
                document.Range().Text = "第一段文本。\n第二段文本。\n第三段文本。";

                if (selection != null)
                {
                    // 检查选择区域类型
                    switch (selection.Type)
                    {
                        case WdSelectionType.wdSelectionIP:
                            Console.WriteLine("插入点选择");
                            break;
                        case WdSelectionType.wdSelectionNormal:
                            Console.WriteLine("正常文本选择");
                            break;
                        case WdSelectionType.wdSelectionColumn:
                            Console.WriteLine("列选择");
                            break;
                        case WdSelectionType.wdSelectionRow:
                            Console.WriteLine("行选择");
                            break;
                        case WdSelectionType.wdSelectionBlock:
                            Console.WriteLine("块选择");
                            break;
                    }

                    // 设置选择模式
                    selection.ExtendMode = false;
                    selection.ColumnSelectMode = false;
                }

                Console.WriteLine("选择区域类型和属性演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"选择区域类型和属性演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 文本选择和操作示例
        /// </summary>
        static void TextSelectionDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;
                var selection = app.Selection;

                // 添加内容
                document.Range().Text = "这是第一段文本。\n这是第二段文本。\n这是第三段文本。";

                if (selection != null)
                {
                    // 选择所有文本
                    selection.WholeStory();

                    // 获取选中的文本
                    string allText = selection.Text;
                    Console.WriteLine($"全部文本: {allText}");

                    // 取消选择
                    selection.Collapse(WdCollapseDirection.wdCollapseEnd);

                    // 选择特定范围
                    selection.SetRange(0, 5);
                    string selectedText = selection.Text;
                    Console.WriteLine($"选中文本: {selectedText}");

                    // 移动选择
                    selection.MoveRight(WdUnits.wdCharacter, 1);
                    selection.MoveDown(WdUnits.wdLine, 1);
                }

                Console.WriteLine("文本选择和操作演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文本选择和操作演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 选择区域的扩展和收缩示例
        /// </summary>
        static void SelectionExpansionDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;
                var selection = app.Selection;

                // 添加内容
                document.Range().Text = "选择区域操作示例文本内容。";

                if (selection != null)
                {
                    // 设置初始位置
                    selection.Collapse(WdCollapseDirection.wdCollapseStart);

                    // 扩展选择区域
                    selection.MoveRight(WdUnits.wdWord, 2, WdMovementType.wdExtend);
                    Console.WriteLine($"扩展后选中文本: '{selection.Text}'");

                    // 进一步扩展
                    selection.MoveRight(WdUnits.wdCharacter, 3, WdMovementType.wdExtend);
                    Console.WriteLine($"再次扩展后选中文本: '{selection.Text}'");

                    // 收缩选择区域
                    selection.MoveLeft(WdUnits.wdWord, 1, WdMovementType.wdExtend);
                    Console.WriteLine($"收缩后选中文本: '{selection.Text}'");
                }

                Console.WriteLine("选择区域扩展和收缩演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"选择区域扩展和收缩演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 高级选择操作示例
        /// </summary>
        static void AdvancedSelectionDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;
                var selection = app.Selection;

                // 添加多段落内容
                document.Range().Text = "第一段内容。\n第二段内容，包含更多文本。\n第三段内容。";

                if (selection != null)
                {
                    // 选择整段
                    selection.EndKey(WdUnits.wdStory);
                    selection.HomeKey(WdUnits.wdLine, WdMovementType.wdExtend);
                    Console.WriteLine($"整段文本: '{selection.Text}'");

                    // 移动到文档开始
                    selection.HomeKey(WdUnits.wdStory);

                    // 选择到文档末尾
                    selection.EndKey(WdUnits.wdStory, WdMovementType.wdExtend);
                    Console.WriteLine($"全文本长度: {selection.Text.Length}");

                    // 取消选择
                    selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                }

                Console.WriteLine("高级选择操作演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"高级选择操作演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 选择区域格式化示例
        /// </summary>
        static void SelectionFormattingDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;
                var selection = app.Selection;

                // 添加内容
                document.Range().Text = "选择区域格式化示例文本。";

                if (selection != null)
                {
                    // 选择所有文本
                    selection.WholeStory();

                    // 应用格式化
                    selection.Font.Bold = true;
                    selection.Font.Italic = true;
                    selection.Font.Underline = WdUnderline.wdUnderlineSingle;
                    selection.Font.Size = 14;
                    selection.Font.Name = "微软雅黑";

                    // 设置段落格式
                    selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    selection.ParagraphFormat.LineSpacing = 1.5f;

                    Console.WriteLine("选择区域格式化演示完成");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"选择区域格式化演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 使用辅助类的完整示例
        /// </summary>
        static void CompleteExampleWithHelpers()
        {
            try
            {
                Console.WriteLine("使用SelectionOperationsManager辅助类进行完整操作:");

                // 创建选择操作管理器实例
                var selectionManager = new SelectionOperationsManager();

                // 执行文本选择和操作
                var selectionResult = selectionManager.PerformTextSelection();
                Console.WriteLine($"文本选择操作结果:");
                Console.WriteLine($"  选择的文本: '{selectionResult.SelectedText}'");
                Console.WriteLine($"  选择类型: {selectionResult.SelectionType}");

                // 执行格式化操作
                var formattingResult = selectionManager.PerformFormattingOperations();
                Console.WriteLine($"格式化操作结果:");
                Console.WriteLine($"  格式化前文本: '{formattingResult.BeforeText}'");
                Console.WriteLine($"  格式化后文本: '{formattingResult.AfterText}'");

                // 执行高级选择操作
                var advancedResult = selectionManager.PerformAdvancedSelection();
                Console.WriteLine($"高级选择操作结果:");
                Console.WriteLine($"  文档段落数: {advancedResult.ParagraphCount}");
                Console.WriteLine($"  选中字符数: {advancedResult.SelectedCharacterCount}");

                Console.WriteLine("使用辅助类的完整示例操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用辅助类的完整示例操作出错: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 选择操作管理器辅助类
    /// </summary>
    public class SelectionOperationsManager
    {
        /// <summary>
        /// 文本选择结果类
        /// </summary>
        public class TextSelectionResult
        {
            /// <summary>
            /// 选择的文本
            /// </summary>
            public string SelectedText { get; set; }

            /// <summary>
            /// 选择类型
            /// </summary>
            public WdSelectionType SelectionType { get; set; }
        }

        /// <summary>
        /// 格式化操作结果类
        /// </summary>
        public class FormattingResult
        {
            /// <summary>
            /// 格式化前文本
            /// </summary>
            public string BeforeText { get; set; }

            /// <summary>
            /// 格式化后文本
            /// </summary>
            public string AfterText { get; set; }
        }

        /// <summary>
        /// 高级选择操作结果类
        /// </summary>
        public class AdvancedSelectionResult
        {
            /// <summary>
            /// 文档段落数
            /// </summary>
            public int ParagraphCount { get; set; }

            /// <summary>
            /// 选中字符数
            /// </summary>
            public int SelectedCharacterCount { get; set; }
        }

        /// <summary>
        /// 执行文本选择操作
        /// </summary>
        /// <returns>文本选择结果</returns>
        public TextSelectionResult PerformTextSelection()
        {
            using var app = WordFactory.BlankWorkbook();
            var document = app.ActiveDocument;
            var selection = app.Selection;

            // 添加内容
            document.Range().Text = "文本选择操作示例内容。";

            if (selection != null)
            {
                // 选择所有文本
                selection.WholeStory();

                return new TextSelectionResult
                {
                    SelectedText = selection.Text,
                    SelectionType = selection.Type
                };
            }

            return new TextSelectionResult();
        }

        /// <summary>
        /// 执行格式化操作
        /// </summary>
        /// <returns>格式化操作结果</returns>
        public FormattingResult PerformFormattingOperations()
        {
            using var app = WordFactory.BlankWorkbook();
            var document = app.ActiveDocument;
            var selection = app.Selection;

            // 添加内容
            document.Range().Text = "格式化操作示例文本内容。";
            string beforeText = document.Range().Text;

            if (selection != null)
            {
                // 选择所有文本
                selection.WholeStory();

                // 应用格式化
                selection.Font.Bold = true;
                selection.Font.Italic = true;
                selection.Font.Name = "楷体";
            }

            return new FormattingResult
            {
                BeforeText = beforeText,
                AfterText = document.Range().Text
            };
        }

        /// <summary>
        /// 执行高级选择操作
        /// </summary>
        /// <returns>高级选择操作结果</returns>
        public AdvancedSelectionResult PerformAdvancedSelection()
        {
            using var app = WordFactory.BlankWorkbook();
            var document = app.ActiveDocument;
            var selection = app.Selection;

            // 添加多段落内容
            document.Range().Text = "第一段内容。\n第二段内容。\n第三段内容。\n第四段内容。";

            int paragraphCount = document.Paragraphs.Count;
            int selectedCharacterCount = 0;

            if (selection != null)
            {
                // 选择前两段
                selection.HomeKey(WdUnits.wdStory);
                for (int i = 0; i < 2; i++)
                {
                    selection.MoveDown(WdUnits.wdLine, 1, WdMovementType.wdExtend);
                }

                selectedCharacterCount = selection.Text.Length;
            }

            return new AdvancedSelectionResult
            {
                ParagraphCount = paragraphCount,
                SelectedCharacterCount = selectedCharacterCount
            };
        }
    }
}
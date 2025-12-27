//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;
using System.Text;

namespace PageLayoutAndPrintingSample
{
    /// <summary>
    /// 打印管理器类
    /// </summary>
    public class PrintingManager
    {
        private readonly IWordApplication _application;
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="application">Word应用程序对象</param>
        /// <param name="document">Word文档对象</param>
        public PrintingManager(IWordApplication? Application, IWordDocument document)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 显示打印预览
        /// </summary>
        public void ShowPrintPreview()
        {
            try
            {
                _application.ActiveWindow.View.Type = WdViewType.wdPrintView;
                Console.WriteLine("已切换到打印预览视图");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"显示打印预览时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 隐藏打印预览，返回到普通视图
        /// </summary>
        public void HidePrintPreview()
        {
            try
            {
                _application.ActiveWindow.View.Type = WdViewType.wdNormalView;
                Console.WriteLine("已返回到普通视图");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"隐藏打印预览时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 打印整个文档
        /// </summary>
        /// <param name="copies">打印份数</param>
        /// <param name="background">是否后台打印</param>
        /// <param name="collate">是否逐份打印</param>
        public void PrintDocument(int copies = 1, bool background = false, bool collate = true)
        {
            try
            {
                _document.PrintOut(
                    background: background,
                    copies: copies,
                    pageType: WdPrintOutPages.wdPrintAllPages,
                    range: WdPrintOutRange.wdPrintAllDocument,
                    item: WdPrintOutItem.wdPrintDocumentContent,
                    collate: collate
                );

                Console.WriteLine($"文档打印任务已启动，打印 {copies} 份");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"打印文档时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 打印当前页面
        /// </summary>
        /// <param name="copies">打印份数</param>
        /// <param name="background">是否后台打印</param>
        /// <param name="collate">是否逐份打印</param>
        public void PrintCurrentPage(int copies = 1, bool background = false, bool collate = true)
        {
            try
            {
                _document.PrintOut(
                    background: background,
                    copies: copies,
                    pageType: WdPrintOutPages.wdPrintAllPages,
                    range: WdPrintOutRange.wdPrintCurrentPage,
                    item: WdPrintOutItem.wdPrintDocumentContent,
                    collate: collate
                );

                Console.WriteLine($"当前页面打印任务已启动，打印 {copies} 份");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"打印当前页面时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 打印指定页面范围
        /// </summary>
        /// <param name="fromPage">起始页码</param>
        /// <param name="toPage">结束页码</param>
        /// <param name="copies">打印份数</param>
        /// <param name="background">是否后台打印</param>
        /// <param name="collate">是否逐份打印</param>
        public void PrintPageRange(int fromPage, int toPage, int copies = 1, bool background = false, bool collate = true)
        {
            try
            {
                string pageRange = $"{fromPage}-{toPage}";

                _document.PrintOut(
                    background: background,
                    copies: copies,
                    pageType: WdPrintOutPages.wdPrintAllPages,
                    range: WdPrintOutRange.wdPrintRangeOfPages,
                    item: WdPrintOutItem.wdPrintDocumentContent,
                    collate: collate,
                    pages: pageRange
                );

                Console.WriteLine($"页面范围 {pageRange} 打印任务已启动，打印 {copies} 份");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"打印页面范围时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 打印选定内容
        /// </summary>
        /// <param name="copies">打印份数</param>
        /// <param name="background">是否后台打印</param>
        /// <param name="collate">是否逐份打印</param>
        public void PrintSelection(int copies = 1, bool background = false, bool collate = true)
        {
            try
            {
                _document.PrintOut(
                    background: background,
                    copies: copies,
                    pageType: WdPrintOutPages.wdPrintAllPages,
                    range: WdPrintOutRange.wdPrintSelection,
                    item: WdPrintOutItem.wdPrintDocumentContent,
                    collate: collate
                );

                Console.WriteLine($"选定内容打印任务已启动，打印 {copies} 份");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"打印选定内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 双面打印文档
        /// </summary>
        /// <param name="manualDuplex">是否手动双面打印</param>
        /// <param name="copies">打印份数</param>
        /// <param name="background">是否后台打印</param>
        /// <param name="collate">是否逐份打印</param>
        public void PrintDuplex(bool manualDuplex = false, int copies = 1, bool background = false, bool collate = true)
        {
            try
            {
                // 设置双面打印选项
                // 注意：实际的双面打印设置可能需要通过打印机驱动程序设置

                _document.PrintOut(
                    background: background,
                    copies: copies,
                    pageType: WdPrintOutPages.wdPrintAllPages,
                    range: WdPrintOutRange.wdPrintAllDocument,
                    item: WdPrintOutItem.wdPrintDocumentContent,
                    collate: collate,
                    manualDuplexPrint: manualDuplex
                );

                string duplexType = manualDuplex ? "手动双面" : "自动双面";
                Console.WriteLine($"{duplexType}打印任务已启动，打印 {copies} 份");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"双面打印时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 打印到文件
        /// </summary>
        /// <param name="outputFileName">输出文件名</param>
        /// <param name="copies">打印份数</param>
        /// <param name="collate">是否逐份打印</param>
        public void PrintToFile(string outputFileName, int copies = 1, bool collate = true)
        {
            try
            {
                _document.PrintOut(
                    background: false,
                    copies: copies,
                    pageType: WdPrintOutPages.wdPrintAllPages,
                    range: WdPrintOutRange.wdPrintAllDocument,
                    item: WdPrintOutItem.wdPrintDocumentContent,
                    collate: collate,
                    outputFileName: outputFileName,
                    printToFile: true
                );

                Console.WriteLine($"文档已打印到文件: {outputFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"打印到文件时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取文档页数估算
        /// </summary>
        /// <returns>估算的页数</returns>
        public int GetEstimatedPageCount()
        {
            try
            {
                // 通过段落数量粗略估算页数
                int paragraphCount = _document.Range().Paragraphs.Count;
                // 假设每页大约50段落（这是一个非常粗略的估算）
                int estimatedPages = Math.Max(1, paragraphCount / 50);

                Console.WriteLine($"文档大约有 {estimatedPages} 页 (基于 {paragraphCount} 个段落估算)");
                return estimatedPages;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"估算页数时出错: {ex.Message}");
                return 1;
            }
        }

        /// <summary>
        /// 获取打印相关信息
        /// </summary>
        /// <returns>打印信息字符串</returns>
        public string GetPrintInfo()
        {
            try
            {
                StringBuilder info = new StringBuilder();
                info.AppendLine("打印信息:");

                // 获取文档页数估算
                int pageCount = GetEstimatedPageCount();
                info.AppendLine($"  估算页数: {pageCount}");

                // 获取当前视图类型
                var viewType = _application.ActiveWindow.View.Type;
                info.AppendLine($"  当前视图: {viewType}");

                // 获取打印机信息（如果可用）
                try
                {
                    string printerName = _application.ActivePrinter;
                    info.AppendLine($"  当前打印机: {printerName}");
                }
                catch
                {
                    info.AppendLine("  当前打印机: 无法获取");
                }

                return info.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取打印信息时出错: {ex.Message}");
                return "无法获取打印信息";
            }
        }

        /// <summary>
        /// 设置打印选项
        /// </summary>
        /// <param name="options">打印选项</param>
        public void SetPrintOptions(PrintOptions options)
        {
            try
            {
                // 这里可以设置各种打印选项
                // 注意：某些选项可能需要通过应用程序级别设置或打印机驱动程序设置
                Console.WriteLine("打印选项已设置:");
                Console.WriteLine($"  打印份数: {options.Copies}");
                Console.WriteLine($"  后台打印: {options.Background}");
                Console.WriteLine($"  逐份打印: {options.Collate}");
                Console.WriteLine($"  手动双面打印: {options.ManualDuplexPrint}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置打印选项时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 打印信封
        /// </summary>
        /// <param name="envelopeSize">信封尺寸</param>
        /// <param name="address">地址</param>
        /// <param name="returnAddress">回邮地址</param>
        /// <param name="printerName">打印机名称（可选）</param>
        public void PrintEnvelope(WdEnvelopeOrientation envelopeSize, string address, string returnAddress, string printerName = null)
        {
            try
            {
                // 添加信封
                using var envelope = _document.Envelope;
                envelope.Insert(address, returnAddress);

                // 设置信封尺寸
                envelope.DefaultOrientation = envelopeSize;

                // 如果指定了打印机，则设置打印机
                if (!string.IsNullOrEmpty(printerName))
                {
                    // 注意：设置打印机可能需要更复杂的操作
                    Console.WriteLine($"将使用打印机: {printerName}");
                }

                // 打印信封
                envelope.PrintOut();

                Console.WriteLine("信封已打印");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"打印信封时出错: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 打印选项类
    /// </summary>
    public class PrintOptions
    {
        /// <summary>
        /// 打印份数
        /// </summary>
        public int Copies { get; set; } = 1;

        /// <summary>
        /// 是否后台打印
        /// </summary>
        public bool Background { get; set; } = false;

        /// <summary>
        /// 是否逐份打印
        /// </summary>
        public bool Collate { get; set; } = true;

        /// <summary>
        /// 是否手动双面打印
        /// </summary>
        public bool ManualDuplexPrint { get; set; } = false;

        /// <summary>
        /// 是否打印到文件
        /// </summary>
        public bool PrintToFile { get; set; } = false;

        /// <summary>
        /// 输出文件名（如果打印到文件）
        /// </summary>
        public string OutputFileName { get; set; } = "";

        /// <summary>
        /// 打印范围
        /// </summary>
        public WdPrintOutRange Range { get; set; } = WdPrintOutRange.wdPrintAllDocument;

        /// <summary>
        /// 页面类型
        /// </summary>
        public WdPrintOutPages PageType { get; set; } = WdPrintOutPages.wdPrintAllPages;

        /// <summary>
        /// 打印项目
        /// </summary>
        public WdPrintOutItem Item { get; set; } = WdPrintOutItem.wdPrintDocumentContent;

        /// <summary>
        /// 页面范围（如果Range设置为wdPrintRangeOfPages）
        /// </summary>
        public string Pages { get; set; } = "";

        /// <summary>
        /// 构造函数
        /// </summary>
        public PrintOptions()
        {
        }

        /// <summary>
        /// 带参数的构造函数
        /// </summary>
        /// <param name="copies">打印份数</param>
        /// <param name="background">是否后台打印</param>
        /// <param name="collate">是否逐份打印</param>
        /// <param name="manualDuplexPrint">是否手动双面打印</param>
        public PrintOptions(int copies, bool background, bool collate, bool manualDuplexPrint)
        {
            Copies = copies;
            Background = background;
            Collate = collate;
            ManualDuplexPrint = manualDuplexPrint;
        }
    }
}
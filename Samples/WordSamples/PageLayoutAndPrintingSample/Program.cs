using MudTools.OfficeInterop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace PageLayoutAndPrintingSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 页面布局和打印示例");

            // 示例1: 页面设置
            Console.WriteLine("\n=== 示例1: 页面设置 ===");
            PageSetupDemo();

            // 示例2: 页眉和页脚
            Console.WriteLine("\n=== 示例2: 页眉和页脚 ===");
            HeaderFooterDemo();

            // 示例3: 分节符和分页符
            Console.WriteLine("\n=== 示例3: 分节符和分页符 ===");
            SectionAndPageBreaksDemo();

            // 示例4: 打印选项和预览
            Console.WriteLine("\n=== 示例4: 打印选项和预览 ===");
            PrintingOptionsDemo();

            // 示例5: 实际应用示例
            Console.WriteLine("\n=== 示例5: 实际应用示例 ===");
            RealWorldPageLayoutDemo();

            // 示例6: 使用辅助类构建完整文档
            Console.WriteLine("\n=== 示例6: 使用辅助类构建完整文档 ===");
            CompleteDocumentWithHelpersDemo();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 页面设置示例
        /// </summary>
        static void PageSetupDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 获取页面设置对象
                var pageSetup = document.Sections[1].PageSetup;

                // 设置纸张大小
                pageSetup.PageWidth = 12240; // A4纸宽度 (单位: 磅/72英寸)
                pageSetup.PageHeight = 15840; // A4纸高度

                // 或者使用预定义的纸张大小
                pageSetup.PageSize = WdPaperSize.wdPaperA4;

                // 设置页面方向
                pageSetup.Orientation = WdOrientation.wdOrientPortrait; // 纵向

                // 设置页边距
                pageSetup.TopMargin = 1440;    // 1英寸 = 72磅
                pageSetup.BottomMargin = 1440;
                pageSetup.LeftMargin = 1800;   // 1.25英寸
                pageSetup.RightMargin = 1800;
                pageSetup.HeaderDistance = 720; // 页眉距离
                pageSetup.FooterDistance = 720; // 页脚距离

                // 设置页面垂直对齐方式
                pageSetup.VerticalAlignment = WdVerticalAlignment.wdAlignVerticalTop;

                // 设置行号
                pageSetup.LineNumbering.Active = 1; // 启用行号
                pageSetup.LineNumbering.RestartMode = WdNumberingRule.wdRestartContinuous;

                Console.WriteLine("页面设置完成");
                Console.WriteLine($"页面宽度: {pageSetup.PageWidth} 磅");
                Console.WriteLine($"页面高度: {pageSetup.PageHeight} 磅");
                Console.WriteLine($"页面方向: {pageSetup.Orientation}");
                Console.WriteLine($"上边距: {pageSetup.TopMargin} 磅");
                Console.WriteLine($"左边距: {pageSetup.LeftMargin} 磅");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"页面设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 页眉和页脚示例
        /// </summary>
        static void HeaderFooterDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 获取页眉和页脚范围
                var headerRange = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                var footerRange = document.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                // 设置页眉内容
                headerRange.Text = "文档标题";
                headerRange.Font.Name = "微软雅黑";
                headerRange.Font.Size = 12;
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 设置页脚内容（包含页码）
                footerRange.Fields.Add(footerRange, WdFieldType.wdFieldPage); // 插入页码
                footerRange.Text = " 第 ";
                footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                footerRange.Fields.Add(footerRange, WdFieldType.wdFieldNumPages); // 插入总页数
                footerRange.Text = " 页";
                footerRange.Font.Size = 10;
                footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 设置首页不同
                document.Sections[1].PageSetup.DifferentFirstPageHeaderFooter = 1;

                // 设置奇偶页不同
                document.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter = 1;

                Console.WriteLine("页眉和页脚设置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"页眉和页脚设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 分节符和分页符示例
        /// </summary>
        static void SectionAndPageBreaksDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加内容
                var range = document.Range();
                range.Text = "第一部分内容\n";

                // 插入分页符
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.InsertBreak(WdBreakType.wdPageBreak);

                // 添加第二部分内容
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "第二部分内容\n";

                // 插入分节符（下一页）
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.InsertBreak(WdBreakType.wdSectionBreakNextPage);

                // 添加第三部分内容
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "第三部分内容\n";

                // 为不同节设置不同的页面布局
                var section1 = document.Sections[1]; // 第一节
                section1.PageSetup.Orientation = WdOrientation.wdOrientPortrait;

                var section2 = document.Sections[2]; // 第二节
                section2.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

                var section3 = document.Sections[3]; // 第三节
                section3.PageSetup.Orientation = WdOrientation.wdOrientPortrait;

                Console.WriteLine("分节符和分页符设置完成");
                Console.WriteLine($"文档共有 {document.Sections.Count} 个节");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"分节符和分页符设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 打印选项和预览示例
        /// </summary>
        static void PrintingOptionsDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 打印预览
                app.ActiveWindow.View.Type = WdViewType.wdPrintPreviewView;
                Console.WriteLine("已切换到打印预览视图");

                // 模拟设置打印选项（不实际打印）
                Console.WriteLine("打印选项设置:");
                Console.WriteLine("  - 打印份数: 2");
                Console.WriteLine("  - 打印所有页面");
                Console.WriteLine("  - 打印文档内容");
                Console.WriteLine("  - 逐份打印: 是");

                // 获取打印相关信息
                int pagesCount = document.Range().Paragraphs.Count; // 粗略估算页数
                Console.WriteLine($"文档大约有 {Math.Max(1, pagesCount / 50)} 页");

                // 返回普通视图
                app.ActiveWindow.View.Type = WdViewType.wdNormalView;
                Console.WriteLine("已返回到普通视图");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"打印选项和预览出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 实际应用示例
        /// </summary>
        static void RealWorldPageLayoutDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                app.Visible = false; // 在实际应用示例中隐藏Word窗口

                var document = app.ActiveDocument;

                // 设置文档属性
                document.Title = "专业文档示例";
                document.Author = "MudTools.OfficeInterop.Word 用户";

                // 设置第一页的页面布局
                var section1 = document.Sections[1];
                var pageSetup = section1.PageSetup;

                // 设置A4纸张
                pageSetup.PageSize = WdPaperSize.wdPaperA4;
                pageSetup.Orientation = WdOrientation.wdOrientPortrait;

                // 设置页边距
                pageSetup.TopMargin = 1440;    // 2厘米
                pageSetup.BottomMargin = 1440;
                pageSetup.LeftMargin = 1800;   // 2.5厘米
                pageSetup.RightMargin = 1800;
                pageSetup.HeaderDistance = 720;
                pageSetup.FooterDistance = 720;

                // 设置首页不同
                pageSetup.DifferentFirstPageHeaderFooter = 1;

                // 添加封面内容
                var coverRange = document.Range();
                coverRange.Text = "\n\n\n";
                coverRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 添加标题
                var titleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                titleRange.Text = "公司年度报告\n";
                titleRange.Font.Name = "微软雅黑";
                titleRange.Font.Size = 28;
                titleRange.Font.Bold = 1;
                titleRange.Font.Color = WdColor.wdColorDarkBlue;
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                titleRange.ParagraphFormat.SpaceAfter = 24;

                // 添加副标题
                var subtitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                subtitleRange.Text = "2025财年总结\n\n\n";
                subtitleRange.Font.Name = "微软雅黑";
                subtitleRange.Font.Size = 18;
                subtitleRange.Font.Color = WdColor.wdColorBlue;
                subtitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 添加公司信息
                var companyRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                companyRange.Text = "某某公司\n";
                companyRange.Font.Name = "宋体";
                companyRange.Font.Size = 14;
                companyRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                var dateRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                dateRange.Text = DateTime.Now.ToString("yyyy年MM月dd日") + "\n";
                dateRange.Font.Name = "宋体";
                dateRange.Font.Size = 12;
                dateRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 插入分页符
                var breakRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                breakRange.InsertBreak(WdBreakType.wdPageBreak);

                // 设置目录页的页眉页脚
                var headerRange = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Text = "公司年度报告";
                headerRange.Font.Name = "宋体";
                headerRange.Font.Size = 10;
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                var footerRange = document.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Text = "第 ";
                footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                footerRange.Fields.Add(footerRange, WdFieldType.wdFieldPage);
                footerRange.Text = " 页";
                footerRange.Font.Size = 10;
                footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 添加目录标题
                var tocTitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                tocTitleRange.Text = "目录\n";
                tocTitleRange.Font.Name = "微软雅黑";
                tocTitleRange.Font.Size = 16;
                tocTitleRange.Font.Bold = 1;
                tocTitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                tocTitleRange.ParagraphFormat.SpaceAfter = 24;

                // 插入分页符
                var breakRange2 = document.Range(document.Content.End - 1, document.Content.End - 1);
                breakRange2.InsertBreak(WdBreakType.wdPageBreak);

                // 添加正文内容
                var contentTitle = document.Range(document.Content.End - 1, document.Content.End - 1);
                contentTitle.Text = "第一章：公司概况\n";
                contentTitle.Font.Name = "微软雅黑";
                contentTitle.Font.Size = 14;
                contentTitle.Font.Bold = 1;
                contentTitle.ParagraphFormat.SpaceAfter = 12;

                var contentRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                contentRange.Text = "这里是公司概况的内容...\n\n";
                contentRange.Font.Name = "宋体";
                contentRange.Font.Size = 12;

                // 添加第二章
                var chapter2Title = document.Range(document.Content.End - 1, document.Content.End - 1);
                chapter2Title.Text = "第二章：财务分析\n";
                chapter2Title.Font.Name = "微软雅黑";
                chapter2Title.Font.Size = 14;
                chapter2Title.Font.Bold = 1;
                chapter2Title.ParagraphFormat.SpaceAfter = 12;

                var chapter2Range = document.Range(document.Content.End - 1, document.Content.End - 1);
                chapter2Range.Text = "这里是财务分析的内容...\n\n";
                chapter2Range.Font.Name = "宋体";
                chapter2Range.Font.Size = 12;

                // 插入分节符（下一页）
                var sectionBreakRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                sectionBreakRange.InsertBreak(WdBreakType.wdSectionBreakNextPage);

                // 为新节设置横向页面
                var section2 = document.Sections[2];
                section2.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

                // 添加横向页面内容
                var landscapeTitle = document.Range(document.Content.End - 1, document.Content.End - 1);
                landscapeTitle.Text = "财务数据图表\n";
                landscapeTitle.Font.Name = "微软雅黑";
                landscapeTitle.Font.Size = 14;
                landscapeTitle.Font.Bold = 1;
                landscapeTitle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                landscapeTitle.ParagraphFormat.SpaceAfter = 12;

                // 保存文档
                string filePath = Path.Combine(Path.GetTempPath(), "PageLayoutDemo.docx");
                document.SaveAs2(filePath);

                Console.WriteLine($"专业文档已创建: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建文档时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 使用辅助类构建完整文档示例
        /// </summary>
        static void CompleteDocumentWithHelpersDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                app.Visible = false; // 隐藏Word窗口

                var document = app.ActiveDocument;

                // 创建文档布局构建器
                var layoutBuilder = new DocumentLayoutBuilder(document);

                // 创建专业文档布局
                layoutBuilder.CreateProfessionalLayout("公司年度报告", "MudTools.OfficeInterop.Word 用户");

                // 添加封面
                layoutBuilder.AddCoverPage(
                    "公司年度报告",
                    "2025财年总结",
                    "某某公司",
                    DateTime.Now
                );

                // 添加目录页
                layoutBuilder.AddTableOfContentsPage();

                // 添加章节内容
                layoutBuilder.AddChapter(
                    "公司概况",
                    "这里是公司概况的详细内容。公司成立于2010年，专注于提供高质量的软件解决方案。我们的团队由经验丰富的开发人员、设计师和项目经理组成，致力于为客户提供卓越的服务。"
                );

                layoutBuilder.AddChapter(
                    "财务分析",
                    "本章详细分析了公司在2025财年的财务表现。总收入达到1.2亿元，同比增长15%。净利润为3000万元，同比增长20%。这些成果反映了公司在市场拓展和成本控制方面的成功。"
                );

                layoutBuilder.AddChapter(
                    "市场展望",
                    "展望未来，公司计划进一步扩大市场份额，投资新技术研发，并加强与合作伙伴的关系。我们相信，通过持续创新和优质服务，公司将在未来几年实现更加显著的增长。"
                );

                // 添加横向页面节
                layoutBuilder.AddLandscapeSection(
                    "财务数据图表",
                    "此页面用于展示财务数据图表，横向布局提供了更多的空间来呈现详细的数据信息。"
                );

                // 设置首页不同的页眉页脚
                layoutBuilder.SetDifferentFirstPageHeaderFooter(
                    "", // 首页无页眉
                    "公司机密", // 首页页脚
                    "", // 其他页面无页眉
                    "第 页" // 其他页面页脚
                );

                // 获取打印管理器并显示打印信息
                var printingManager = new PrintingManager(app, document);
                string printInfo = printingManager.GetPrintInfo();
                Console.WriteLine(printInfo);

                // 保存文档
                string filePath = Path.Combine(Path.GetTempPath(), "CompleteDocumentWithHelpers.docx");
                document.SaveAs2(filePath);

                Console.WriteLine($"使用辅助类创建的完整文档已保存: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用辅助类创建完整文档时出错: {ex.Message}");
            }
        }
    }
}
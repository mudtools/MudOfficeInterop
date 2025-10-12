using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PageLayoutAndPrintingSample
{
    /// <summary>
    /// 文档布局构建器类
    /// </summary>
    public class DocumentLayoutBuilder
    {
        private readonly IWordDocument _document;
        private readonly PageSetupHelper _pageSetupHelper;
        private readonly HeaderFooterHelper _headerFooterHelper;
        private readonly SectionManager _sectionManager;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public DocumentLayoutBuilder(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _pageSetupHelper = new PageSetupHelper(document);
            _headerFooterHelper = new HeaderFooterHelper(document);
            _sectionManager = new SectionManager(document);
        }

        /// <summary>
        /// 创建专业文档布局
        /// </summary>
        /// <param name="title">文档标题</param>
        /// <param name="author">文档作者</param>
        public void CreateProfessionalLayout(string title, string author)
        {
            try
            {
                // 设置文档属性
                _document.Title = title;
                _document.Author = author;

                // 设置第一页的页面布局
                var section1 = _document.Sections[1];
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

                Console.WriteLine("专业文档布局已创建");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建专业文档布局时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加封面内容
        /// </summary>
        /// <param name="title">标题</param>
        /// <param name="subtitle">副标题</param>
        /// <param name="company">公司名称</param>
        /// <param name="date">日期</param>
        public void AddCoverPage(string title, string subtitle, string company, DateTime? date = null)
        {
            try
            {
                // 添加封面内容
                var coverRange = _document.Range();
                coverRange.Text = "\n\n\n";
                coverRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 添加标题
                var titleRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                titleRange.Text = title + "\n";
                titleRange.Font.Name = "微软雅黑";
                titleRange.Font.Size = 28;
                titleRange.Font.Bold = 1;
                titleRange.Font.Color = WdColor.wdColorDarkBlue;
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                titleRange.ParagraphFormat.SpaceAfter = 24;

                // 添加副标题
                var subtitleRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                subtitleRange.Text = subtitle + "\n\n\n";
                subtitleRange.Font.Name = "微软雅黑";
                subtitleRange.Font.Size = 18;
                subtitleRange.Font.Color = WdColor.wdColorBlue;
                subtitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 添加公司信息
                var companyRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                companyRange.Text = company + "\n";
                companyRange.Font.Name = "宋体";
                companyRange.Font.Size = 14;
                companyRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                var dateRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                dateRange.Text = (date ?? DateTime.Now).ToString("yyyy年MM月dd日") + "\n";
                dateRange.Font.Name = "宋体";
                dateRange.Font.Size = 12;
                dateRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                Console.WriteLine("封面页面已添加");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加封面页面时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加目录页
        /// </summary>
        /// <param name="title">目录标题</param>
        public void AddTableOfContentsPage(string title = "目录")
        {
            try
            {
                // 插入分页符
                var breakRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                _sectionManager.InsertPageBreak(breakRange);

                // 设置目录页的页眉页脚
                _headerFooterHelper.AddHeaderText(
                    1,
                    WdHeaderFooterIndex.wdHeaderFooterPrimary,
                    _document.Title ?? "文档标题",
                    "宋体",
                    10,
                    WdParagraphAlignment.wdAlignParagraphCenter
                );

                _headerFooterHelper.AddPageNumber(
                    1,
                    WdHeaderFooterIndex.wdHeaderFooterPrimary,
                    "第 ",
                    " 页",
                    "宋体",
                    10,
                    WdParagraphAlignment.wdAlignParagraphCenter
                );

                // 添加目录标题
                var tocTitleRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                tocTitleRange.Text = title + "\n";
                tocTitleRange.Font.Name = "微软雅黑";
                tocTitleRange.Font.Size = 16;
                tocTitleRange.Font.Bold = 1;
                tocTitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                tocTitleRange.ParagraphFormat.SpaceAfter = 24;

                Console.WriteLine("目录页面已添加");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加目录页面时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加章节内容
        /// </summary>
        /// <param name="chapterTitle">章节标题</param>
        /// <param name="content">章节内容</param>
        /// <param name="chapterNumber">章节编号</param>
        public void AddChapter(string chapterTitle, string content, int chapterNumber = 1)
        {
            try
            {
                // 插入分页符
                var breakRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                _sectionManager.InsertPageBreak(breakRange);

                // 添加章节标题
                var chapterTitleRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                chapterTitleRange.Text = $"第{chapterNumber}章：{chapterTitle}\n";
                chapterTitleRange.Font.Name = "微软雅黑";
                chapterTitleRange.Font.Size = 14;
                chapterTitleRange.Font.Bold = 1;
                chapterTitleRange.ParagraphFormat.SpaceAfter = 12;

                // 添加章节内容
                var contentRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                contentRange.Text = content + "\n\n";
                contentRange.Font.Name = "宋体";
                contentRange.Font.Size = 12;

                Console.WriteLine($"第{chapterNumber}章已添加");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加章节时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加横向页面节
        /// </summary>
        /// <param name="title">页面标题</param>
        /// <param name="content">页面内容</param>
        public void AddLandscapeSection(string title, string content)
        {
            try
            {
                // 插入分节符（下一页）
                _sectionManager.InsertSectionBreakNextPageAtEnd();

                // 为新节设置横向页面
                int sectionCount = _document.Sections.Count;
                _pageSetupHelper.SetPageOrientation(sectionCount, WdOrientation.wdOrientLandscape);

                // 添加横向页面内容
                var landscapeTitle = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                landscapeTitle.Text = title + "\n";
                landscapeTitle.Font.Name = "微软雅黑";
                landscapeTitle.Font.Size = 14;
                landscapeTitle.Font.Bold = 1;
                landscapeTitle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                landscapeTitle.ParagraphFormat.SpaceAfter = 12;

                var landscapeContent = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                landscapeContent.Text = content + "\n";
                landscapeContent.Font.Name = "宋体";
                landscapeContent.Font.Size = 12;

                Console.WriteLine("横向页面节已添加");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加横向页面节时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 设置首页不同的页眉页脚
        /// </summary>
        /// <param name="firstPageHeaderText">首页页眉文本</param>
        /// <param name="firstPageFooterText">首页页脚文本</param>
        /// <param name="otherPagesHeaderText">其他页面页眉文本</param>
        /// <param name="otherPagesFooterText">其他页面页脚文本</param>
        public void SetDifferentFirstPageHeaderFooter(
            string firstPageHeaderText,
            string firstPageFooterText,
            string otherPagesHeaderText,
            string otherPagesFooterText)
        {
            try
            {
                _headerFooterHelper.SetDifferentFirstPageHeaderFooter(
                    1,
                    firstPageHeaderText,
                    firstPageFooterText,
                    otherPagesHeaderText,
                    otherPagesFooterText
                );
                
                Console.WriteLine("首页不同的页眉页脚已设置");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置首页不同的页眉页脚时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 设置奇偶页不同的页眉页脚
        /// </summary>
        /// <param name="oddHeaderText">奇数页页眉文本</param>
        /// <param name="oddFooterText">奇数页页脚文本</param>
        /// <param name="evenHeaderText">偶数页页眉文本</param>
        /// <param name="evenFooterText">偶数页页脚文本</param>
        public void SetOddAndEvenPagesHeaderFooter(
            string oddHeaderText,
            string oddFooterText,
            string evenHeaderText,
            string evenFooterText)
        {
            try
            {
                _headerFooterHelper.SetOddAndEvenPagesHeaderFooter(
                    1,
                    oddHeaderText,
                    oddFooterText,
                    evenHeaderText,
                    evenFooterText
                );
                
                Console.WriteLine("奇偶页不同的页眉页脚已设置");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置奇偶页不同的页眉页脚时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取页面设置助手
        /// </summary>
        public PageSetupHelper PageSetupHelper => _pageSetupHelper;

        /// <summary>
        /// 获取页眉页脚助手
        /// </summary>
        public HeaderFooterHelper HeaderFooterHelper => _headerFooterHelper;

        /// <summary>
        /// 获取节管理器
        /// </summary>
        public SectionManager SectionManager => _sectionManager;
    }
}
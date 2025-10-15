//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace ReportGenerationSystemSample
{
    /// <summary>
    /// 报表模板管理器类
    /// </summary>
    public class ReportTemplateManager
    {
        private readonly IWordApplication _application;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="application">Word应用程序对象</param>
        public ReportTemplateManager(IWordApplication application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        /// <summary>
        /// 创建销售报表模板
        /// </summary>
        /// <param name="templatePath">模板保存路径</param>
        /// <returns>是否创建成功</returns>
        public bool CreateSalesReportTemplate(string templatePath)
        {
            try
            {
                var document = _application.ActiveDocument;

                // 设置文档属性
                document.Title = "销售报表模板";

                // 添加页眉
                var headerRange = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Text = "公司月度销售报表";
                headerRange.Font.Name = "微软雅黑";
                headerRange.Font.Size = 14;
                headerRange.Font.Bold = true;
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 添加页脚（包含页码）
                var footerRange = document.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Text = "第 ";
                footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                footerRange.Fields.Add(footerRange, WdFieldType.wdFieldPage);
                footerRange.Text = " 页";
                footerRange.Font.Size = 10;
                footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 添加标题
                var titleRange = document.Range();
                titleRange.Text = "XYZ公司月度销售报表\n";
                titleRange.Font.Name = "微软雅黑";
                titleRange.Font.Size = 20;
                titleRange.Font.Bold = true;
                titleRange.Font.Color = WdColor.wdColorDarkBlue;
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                titleRange.ParagraphFormat.SpaceAfter = 24;

                // 添加报表信息
                var infoRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                infoRange.Text = "报表期间：{REPORT_PERIOD}\n";
                infoRange.Font.Name = "宋体";
                infoRange.Font.Size = 12;

                infoRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                infoRange.Text = "生成时间：{GENERATION_TIME}\n";
                infoRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                infoRange.Text = "报表类型：{REPORT_TYPE}\n\n";

                // 添加数据表格标题
                var tableTitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                tableTitleRange.Text = "销售数据详情\n";
                tableTitleRange.Font.Name = "微软雅黑";
                tableTitleRange.Font.Size = 16;
                tableTitleRange.Font.Bold = true;
                tableTitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                tableTitleRange.ParagraphFormat.SpaceAfter = 12;

                // 创建表格占位符
                var tableRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                tableRange.Text = "{SALES_DATA_TABLE}";

                // 添加图表占位符
                var chartRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                chartRange.Text = "\n\n{SALES_CHART}";

                // 添加总结部分
                var summaryTitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                summaryTitleRange.Text = "\n\n销售总结\n";
                summaryTitleRange.Font.Size = 16;
                summaryTitleRange.Font.Bold = true;
                summaryTitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                var summaryRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                summaryRange.Text = "\n总销售额：{TOTAL_SALES}\n";
                summaryRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                summaryRange.Text = "同比增长：{YEAR_OVER_YEAR_GROWTH}\n";
                summaryRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                summaryRange.Text = "环比增长：{MONTH_OVER_MONTH_GROWTH}\n";

                // 保存模板
                document.SaveAs(templatePath);

                Console.WriteLine($"销售报表模板已创建: {templatePath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建销售报表模板时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建财务报表模板
        /// </summary>
        /// <param name="templatePath">模板保存路径</param>
        /// <returns>是否创建成功</returns>
        public bool CreateFinancialReportTemplate(string templatePath)
        {
            try
            {
                var document = _application.ActiveDocument;

                // 设置文档属性
                document.Title = "财务报表模板";

                // 添加页眉
                var headerRange = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Text = "公司月度财务报表";
                headerRange.Font.Name = "微软雅黑";
                headerRange.Font.Size = 14;
                headerRange.Font.Bold = true;
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 添加页脚
                var footerRange = document.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Text = "第 ";
                footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                footerRange.Fields.Add(footerRange, WdFieldType.wdFieldPage);
                footerRange.Text = " 页";
                footerRange.Font.Size = 10;
                footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 添加标题
                var titleRange = document.Range();
                titleRange.Text = "XYZ公司月度财务报表\n";
                titleRange.Font.Name = "微软雅黑";
                titleRange.Font.Size = 20;
                titleRange.Font.Bold = true;
                titleRange.Font.Color = WdColor.wdColorDarkBlue;
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                titleRange.ParagraphFormat.SpaceAfter = 24;

                // 添加报表信息
                var infoRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                infoRange.Text = "报表期间：{REPORT_PERIOD}\n";
                infoRange.Font.Name = "宋体";
                infoRange.Font.Size = 12;

                infoRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                infoRange.Text = "生成时间：{GENERATION_TIME}\n";
                infoRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                infoRange.Text = "报表类型：{REPORT_TYPE}\n\n";

                // 添加收入部分
                var incomeTitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                incomeTitleRange.Text = "收入明细\n";
                incomeTitleRange.Font.Name = "微软雅黑";
                incomeTitleRange.Font.Size = 16;
                incomeTitleRange.Font.Bold = true;
                incomeTitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                incomeTitleRange.ParagraphFormat.SpaceAfter = 12;

                var incomeRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                incomeRange.Text = "{INCOME_TABLE}";

                // 添加支出部分
                var expenseTitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                expenseTitleRange.Text = "\n\n支出明细\n";
                expenseTitleRange.Font.Name = "微软雅黑";
                expenseTitleRange.Font.Size = 16;
                expenseTitleRange.Font.Bold = true;
                expenseTitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                expenseTitleRange.ParagraphFormat.SpaceAfter = 12;

                var expenseRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                expenseRange.Text = "{EXPENSE_TABLE}";

                // 添加总结部分
                var summaryTitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                summaryTitleRange.Text = "\n\n财务总结\n";
                summaryTitleRange.Font.Size = 16;
                summaryTitleRange.Font.Bold = true;
                summaryTitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                var summaryRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                summaryRange.Text = "\n总收入：{TOTAL_INCOME}\n";
                summaryRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                summaryRange.Text = "总支出：{TOTAL_EXPENSE}\n";
                summaryRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                summaryRange.Text = "净利润：{NET_PROFIT}\n";
                summaryRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                summaryRange.Text = "利润率：{PROFIT_MARGIN}\n";

                // 保存模板
                document.SaveAs(templatePath);

                Console.WriteLine($"财务报表模板已创建: {templatePath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建财务报表模板时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建项目进度报表模板
        /// </summary>
        /// <param name="templatePath">模板保存路径</param>
        /// <returns>是否创建成功</returns>
        public bool CreateProjectProgressReportTemplate(string templatePath)
        {
            try
            {
                var document = _application.ActiveDocument;

                // 设置文档属性
                document.Title = "项目进度报表模板";

                // 添加页眉
                var headerRange = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Text = "项目进度报表";
                headerRange.Font.Name = "微软雅黑";
                headerRange.Font.Size = 14;
                headerRange.Font.Bold = true;
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 添加页脚
                var footerRange = document.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Text = "第 ";
                footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                footerRange.Fields.Add(footerRange, WdFieldType.wdFieldPage);
                footerRange.Text = " 页";
                footerRange.Font.Size = 10;
                footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 添加标题
                var titleRange = document.Range();
                titleRange.Text = "XYZ公司项目进度报表\n";
                titleRange.Font.Name = "微软雅黑";
                titleRange.Font.Size = 20;
                titleRange.Font.Bold = true;
                titleRange.Font.Color = WdColor.wdColorDarkBlue;
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                titleRange.ParagraphFormat.SpaceAfter = 24;

                // 添加报表信息
                var infoRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                infoRange.Text = "项目名称：{PROJECT_NAME}\n";
                infoRange.Font.Name = "宋体";
                infoRange.Font.Size = 12;

                infoRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                infoRange.Text = "报表期间：{REPORT_PERIOD}\n";
                infoRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                infoRange.Text = "生成时间：{GENERATION_TIME}\n\n";

                // 添加项目概述
                var overviewTitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                overviewTitleRange.Text = "项目概述\n";
                overviewTitleRange.Font.Name = "微软雅黑";
                overviewTitleRange.Font.Size = 16;
                overviewTitleRange.Font.Bold = true;
                overviewTitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                overviewTitleRange.ParagraphFormat.SpaceAfter = 12;

                var overviewRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                overviewRange.Text = "{PROJECT_OVERVIEW}";

                // 添加进度详情
                var progressTitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                progressTitleRange.Text = "\n\n进度详情\n";
                progressTitleRange.Font.Name = "微软雅黑";
                progressTitleRange.Font.Size = 16;
                progressTitleRange.Font.Bold = true;
                progressTitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                progressTitleRange.ParagraphFormat.SpaceAfter = 12;

                var progressRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                progressRange.Text = "{PROGRESS_DETAILS}";

                // 添加里程碑
                var milestoneTitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                milestoneTitleRange.Text = "\n\n里程碑\n";
                milestoneTitleRange.Font.Name = "微软雅黑";
                milestoneTitleRange.Font.Size = 16;
                milestoneTitleRange.Font.Bold = true;
                milestoneTitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                milestoneTitleRange.ParagraphFormat.SpaceAfter = 12;

                var milestoneRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                milestoneRange.Text = "{MILESTONES}";

                // 添加风险和问题
                var riskTitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                riskTitleRange.Text = "\n\n风险和问题\n";
                riskTitleRange.Font.Name = "微软雅黑";
                riskTitleRange.Font.Size = 16;
                riskTitleRange.Font.Bold = true;
                riskTitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                riskTitleRange.ParagraphFormat.SpaceAfter = 12;

                var riskRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                riskRange.Text = "{RISKS_AND_ISSUES}";

                // 添加总结
                var summaryTitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                summaryTitleRange.Text = "\n\n总结\n";
                summaryTitleRange.Font.Size = 16;
                summaryTitleRange.Font.Bold = true;
                summaryTitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                var summaryRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                summaryRange.Text = "\n项目状态：{PROJECT_STATUS}\n";
                summaryRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                summaryRange.Text = "预计完成时间：{ESTIMATED_COMPLETION}\n";
                summaryRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                summaryRange.Text = "预算使用率：{BUDGET_USAGE}\n";

                // 保存模板
                document.SaveAs(templatePath);

                Console.WriteLine($"项目进度报表模板已创建: {templatePath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建项目进度报表模板时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 加载报表模板
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <returns>Word文档对象</returns>
        public IWordDocument LoadTemplate(string templatePath)
        {
            try
            {
                var app = WordFactory.CreateFrom(templatePath);
                var document = app.ActiveDocument;
                Console.WriteLine($"报表模板已加载: {templatePath}");
                return document;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载报表模板时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 获取模板信息
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <returns>模板信息</returns>
        public TemplateInfo GetTemplateInfo(string templatePath)
        {
            var templateInfo = new TemplateInfo();

            try
            {
                using var app = WordFactory.CreateFrom(templatePath);
                var document = app.ActiveDocument;

                templateInfo.Title = document.Title;
                templateInfo.PageCount = document.Range().Paragraphs.Count;
                templateInfo.Placeholders = ExtractPlaceholders(document);

                Console.WriteLine($"模板信息已获取: {templatePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取模板信息时出错: {ex.Message}");
                templateInfo.ErrorMessage = ex.Message;
            }

            return templateInfo;
        }

        /// <summary>
        /// 提取模板中的占位符
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <returns>占位符列表</returns>
        private List<string> ExtractPlaceholders(IWordDocument document)
        {
            var placeholders = new List<string>();
            try
            {
                var text = document.Range().Text;
                // 简单提取{XXX}格式的占位符
                var placeholderMatches = System.Text.RegularExpressions.Regex.Matches(text, @"\{[^}]+\}");
                foreach (System.Text.RegularExpressions.Match match in placeholderMatches)
                {
                    if (!placeholders.Contains(match.Value))
                    {
                        placeholders.Add(match.Value);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取占位符时出错: {ex.Message}");
            }

            return placeholders;
        }
    }

    /// <summary>
    /// 模板信息类
    /// </summary>
    public class TemplateInfo
    {
        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 主题
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// 作者
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// 关键词
        /// </summary>
        public string Keywords { get; set; }

        /// <summary>
        /// 页数
        /// </summary>
        public int PageCount { get; set; }

        /// <summary>
        /// 占位符列表
        /// </summary>
        public List<string> Placeholders { get; set; } = new List<string>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成模板信息报告
        /// </summary>
        /// <returns>模板信息报告</returns>
        public string GenerateReport()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"获取模板信息失败: {ErrorMessage}";
            }

            var placeholders = string.Join(", ", Placeholders);

            return $"模板信息报告:\n" +
                   $"  标题: {Title}\n" +
                   $"  主题: {Subject}\n" +
                   $"  作者: {Author}\n" +
                   $"  关键词: {Keywords}\n" +
                   $"  段落数: {PageCount}\n" +
                   $"  占位符: {placeholders}";
        }
    }
}
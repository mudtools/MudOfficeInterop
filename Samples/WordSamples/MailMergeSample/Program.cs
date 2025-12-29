//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace MailMergeSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 邮件合并示例");

            // 示例1: 邮件合并基础
            Console.WriteLine("\n=== 示例1: 邮件合并基础 ===");
            MailMergeBasicsDemo();

            // 示例2: 数据源连接
            Console.WriteLine("\n=== 示例2: 数据源连接 ===");
            DataSourceConnectionDemo();

            // 示例3: 合并字段操作
            Console.WriteLine("\n=== 示例3: 合并字段操作 ===");
            MergeFieldOperationsDemo();

            // 示例4: 执行邮件合并
            Console.WriteLine("\n=== 示例4: 执行邮件合并 ===");
            ExecuteMailMergeDemo();

            // 示例5: 高级邮件合并操作
            Console.WriteLine("\n=== 示例5: 高级邮件合并操作 ===");
            AdvancedMailMergeOperationsDemo();

            // 示例6: 实际应用示例
            Console.WriteLine("\n=== 示例6: 实际应用示例 ===");
            RealWorldApplicationDemo();

            // 示例7: 使用辅助类的完整示例
            Console.WriteLine("\n=== 示例7: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 邮件合并基础示例
        /// </summary>
        static void MailMergeBasicsDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;

                // 设置文档为邮件合并主文档
                using var mailMerge = document.MailMerge;

                // 检查文档是否为邮件合并主文档
                bool isMailMergeDoc = mailMerge.MainDocumentType != WdMailMergeMainDocType.wdNotAMergeDocument;
                Console.WriteLine($"是否为邮件合并文档: {isMailMergeDoc}");

                // 设置主文档类型
                mailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters; // 信函
                Console.WriteLine("已设置主文档类型为信函");

                Console.WriteLine("邮件合并基础演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"邮件合并基础演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 数据源连接示例
        /// </summary>
        static void DataSourceConnectionDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;
                using var mailMerge = document.MailMerge;

                // 连接Excel数据源示例（注释掉实际执行以避免文件依赖）
                /*
                try
                {
                    mailMerge.OpenDataSource(
                        Name: @"C:\data\Customers.xlsx",
                        Connection: "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\data\\Customers.xlsx;Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"",
                        SQLStatement: "SELECT * FROM [Sheet1$]"
                    );

                    Console.WriteLine("数据源连接成功");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"数据源连接失败: {ex.Message}");
                }
                */

                Console.WriteLine("数据源连接演示完成（实际连接已注释）");

                // 连接文本文件数据源示例（注释掉实际执行以避免文件依赖）
                /*
                try
                {
                    mailMerge.OpenDataSource(
                        Name: @"C:\data\Customers.csv",
                        Connection: "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\data\\;Extended Properties=\"text;HDR=YES;FMT=Delimited\"",
                        SQLStatement: "SELECT * FROM Customers.csv"
                    );

                    Console.WriteLine("CSV数据源连接成功");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"CSV数据源连接失败: {ex.Message}");
                }
                */

                Console.WriteLine("CSV数据源连接演示完成（实际连接已注释）");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"数据源连接演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 合并字段操作示例
        /// </summary>
        static void MergeFieldOperationsDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;
                using var mailMerge = document.MailMerge;

                // 添加合并字段到文档
                using var range = document.Range();

                // 添加标题
                range.Text = "客户信息\n\n";
                range.Font.Bold = true;
                range.Font.Size = 16;
                range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 移动到文档末尾
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 添加包含合并字段的内容
                range.Text = "尊敬的 ";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 插入合并字段
                mailMerge.Fields.Add(range, "姓名");
                range.Text = " 先生/女士：\n\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                range.Text = "您的客户编号是：";
                mailMerge.Fields.Add(range, "客户编号");
                range.Text = "\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                range.Text = "联系电话：";
                mailMerge.Fields.Add(range, "电话");
                range.Text = "\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                range.Text = "电子邮箱：";
                mailMerge.Fields.Add(range, "邮箱");
                range.Text = "\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                range.Text = "地址：";
                mailMerge.Fields.Add(range, "地址");
                range.Text = "\n\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                range.Text = "感谢您选择我们的服务！\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 查看所有合并字段
                Console.WriteLine($"合并字段数量: {mailMerge.Fields.Count}");
                for (int i = 1; i <= mailMerge.Fields.Count; i++)
                {
                    Console.WriteLine($"字段 {i}: {mailMerge.Fields[i].Code.Text}");
                }

                Console.WriteLine("合并字段操作演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"合并字段操作演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行邮件合并示例
        /// </summary>
        static void ExecuteMailMergeDemo()
        {
            try
            {
                // 注意：实际执行需要模板文件和数据源
                // 这里仅演示代码结构
                Console.WriteLine("执行邮件合并演示（需要实际模板和数据源）");

                /*
                using var app = WordFactory.CreateFrom(@"C:\templates\LetterTemplate.dotx");
                var document = app.ActiveDocument;
                var mailMerge = document.MailMerge;

                try
                {
                    // 连接数据源
                    mailMerge.OpenDataSource(
                        Name: @"C:\data\Customers.xlsx",
                        Connection: "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\data\\Customers.xlsx;Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"",
                        SQLStatement: "SELECT * FROM [Sheet1$]"
                    );

                    // 设置主文档类型
                    mailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters;

                    // 执行邮件合并到新文档
                    mailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;
                    mailMerge.Execute(Pause: false);

                    Console.WriteLine("邮件合并执行完成");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"邮件合并执行失败: {ex.Message}");
                }
                */

                Console.WriteLine("执行邮件合并演示完成（实际执行已注释）");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"执行邮件合并演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 高级邮件合并操作示例
        /// </summary>
        static void AdvancedMailMergeOperationsDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;
                using var mailMerge = document.MailMerge;

                // 添加条件合并字段
                using var range = document.Range();
                range.Text = "亲爱的客户：\n\n";

                // 添加条件文本（根据性别）
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "{ IF { MERGEFIELD 性别 } = \"男\" \"先生\" \"女士\" }";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n\n";

                // 添加计算字段（计算折扣）
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "您的订单总额为：{ MERGEFIELD 订单金额 } 元\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "享受折扣后价格为：{ = { MERGEFIELD 订单金额 } * 0.9 } 元\n\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 添加日期字段
                range.Text = "生成日期：{ DATE \\@ \"yyyy年MM月dd日\" }\n\n";

                // 更新所有字段
                document.Range().Fields.Update();

                Console.WriteLine("高级邮件合并操作演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"高级邮件合并操作演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 实际应用示例
        /// </summary>
        static void RealWorldApplicationDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                app.Visible = false; // 在实际应用示例中隐藏Word窗口

                using var document = app.ActiveDocument;
                using var mailMerge = document.MailMerge;

                Console.WriteLine("开始创建邮件合并系统...");

                // 1. 创建邮件合并模板
                CreateMailMergeTemplate(document);

                // 2. 保存模板
                string templatePath = Path.Combine(Path.GetTempPath(), "CustomerLetterTemplate.dotx");
                document.Save(templatePath);
                Console.WriteLine($"邮件合并模板已创建: {templatePath}");

                // 3. 演示执行邮件合并
                Console.WriteLine("执行邮件合并演示（需要实际数据源）");

                Console.WriteLine("客户信函生成演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"邮件合并系统演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 创建邮件合并模板
        /// </summary>
        /// <param name="document">文档对象</param>
        private static void CreateMailMergeTemplate(IWordDocument document)
        {
            try
            {
                using var mailMerge = document.MailMerge;

                Console.WriteLine("创建邮件合并模板...");

                // 设置文档为邮件合并主文档
                mailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters;

                // 添加模板内容
                using var range = document.Range();

                // 页眉
                range.Text = "ABC有限公司\n地址：某某市某某区某某路123号\n电话：010-12345678\n\n";
                range.Font.Bold = true;
                range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 日期
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "{ DATE \\@ \"yyyy年MM月dd日\" }\n\n";
                range.Font.Bold = false;
                range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                // 收件人地址
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "{ MERGEFIELD 客户姓名 }\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "{ MERGEFIELD 地址 }\n\n";

                // 称呼
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "尊敬的 { MERGEFIELD 客户姓名 } 先生/女士：\n\n";

                // 正文
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "感谢您一直以来对我们公司的支持与信任。我们很高兴地通知您，您的账户信息已更新。\n\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "以下是您的账户信息：\n\n";

                // 账户信息表格
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "客户编号：\t{ MERGEFIELD 客户编号 }\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "账户余额：\t{ MERGEFIELD 账户余额 } 元\n";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "信用等级：\t{ MERGEFIELD 信用等级 }\n\n";

                // 条款和条件
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "如有任何疑问，请随时与我们联系。\n\n";

                // 结尾
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "此致\n敬礼！\n\nABC有限公司客户服务部\n";

                // 更新字段
                document.Range().Fields.Update();

                Console.WriteLine("邮件合并模板创建完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建邮件合并模板时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 使用辅助类的完整示例
        /// </summary>
        static void CompleteExampleWithHelpers()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                app.Visible = false; // 隐藏Word窗口

                using var document = app.ActiveDocument;

                // 创建邮件合并助手
                var mailMergeHelper = new MailMergeHelper(document);

                // 设置主文档类型
                mailMergeHelper.SetMainDocumentType(WdMailMergeMainDocType.wdFormLetters);

                // 检查是否为邮件合并文档
                bool isMailMergeDoc = mailMergeHelper.IsMailMergeDocument();

                // 创建模板构建器
                var templateBuilder = new MailMergeTemplateBuilder(document);

                // 创建信函模板
                bool templateCreated = templateBuilder.CreateLetterTemplate(
                    "ABC有限公司",
                    "某某市某某区某某路123号",
                    "010-12345678");

                if (templateCreated)
                {
                    Console.WriteLine("信函模板创建成功");

                    // 获取所有合并字段
                    var fields = mailMergeHelper.GetAllMergeFields();
                    Console.WriteLine($"找到 {fields.Count} 个合并字段");

                    // 创建数据源管理器
                    var dataSourceManager = new DataSourceManager(document);

                    // 创建示例客户数据
                    var sampleData = dataSourceManager.CreateSampleCustomerData();
                    Console.WriteLine($"创建了 {sampleData.Count} 条示例客户数据");

                    // 创建执行器
                    var executor = new MailMergeExecutor(app, document);

                    // 检查合并设置
                    var checkResult = executor.CheckMergeSettings();
                    Console.WriteLine(checkResult.GenerateReport());

                    // 获取统计信息
                    var stats = executor.GetMergeStatistics();
                    Console.WriteLine(stats.GenerateReport());

                    // 使用示例数据执行邮件合并
                    var executionResult = executor.ExecuteWithSampleData(
                        sampleData,
                        Path.Combine(Path.GetTempPath(), "CustomerLettersSample.docx"));

                    Console.WriteLine(executionResult.GenerateReport());
                }

                Console.WriteLine("使用辅助类的完整示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用辅助类的完整示例演示出错: {ex.Message}");
            }
        }
    }
}
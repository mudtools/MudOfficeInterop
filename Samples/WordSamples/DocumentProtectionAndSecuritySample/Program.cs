//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace DocumentProtectionAndSecuritySample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 文档保护和安全示例");

            // 示例1: 密码保护
            Console.WriteLine("\n=== 示例1: 密码保护 ===");
            PasswordProtectionDemo();

            // 示例2: 编辑限制
            Console.WriteLine("\n=== 示例2: 编辑限制 ===");
            EditingRestrictionsDemo();

            // 示例3: 内容保护
            Console.WriteLine("\n=== 示例3: 内容保护 ===");
            ContentProtectionDemo();

            // 示例5: 文档权限管理
            Console.WriteLine("\n=== 示例5: 文档权限管理 ===");
            DocumentPermissionManagementDemo();

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
        /// 密码保护示例
        /// </summary>
        static void PasswordProtectionDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加示例内容
                document.Range().Text = "这是受保护的敏感文档内容。\n包含重要的商业信息。";

                // 设置打开密码
                document.Password = "OpenPassword123";
                Console.WriteLine("已设置打开密码");

                // 设置修改密码（编辑密码）
                document.WritePassword = "EditPassword456";
                Console.WriteLine("已设置修改密码");

                // 设置密码加密选项
                document.EncryptionProvider = "Microsoft Enhanced RSA and AES Cryptographic Provider";
                Console.WriteLine("已设置加密提供程序");

                Console.WriteLine("密码保护演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"密码保护演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 编辑限制示例
        /// </summary>
        static void EditingRestrictionsDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加示例内容
                var range = document.Range();
                range.Text = "文档标题\n\n这是可以编辑的内容区域。\n\n这是受保护的内容区域，不能编辑。\n\n这是另一个可编辑区域。";

                // 定义可编辑区域
                var editableRange1 = document.Range(15, 27); // "这是可以编辑的内容区域。"
                var editableRange2 = document.Range(65, 78); // "这是另一个可编辑区域。"

                // 添加可编辑区域
                var ed1 = document.EditableRanges.Add(editableRange1);
                ed1.Editors.Add(WdEditorType.wdEditorEveryone); // 所有人可编辑

                var ed2 = document.EditableRanges.Add(editableRange2);
                ed2.Editors.Add("特定用户组"); // 特定用户组可编辑

                // 应用编辑限制
                document.Protect(
                    protectionType: WdProtectionType.wdAllowOnlyReading, // 只读保护
                    noReset: true,
                    password: "ProtectionPass123");

                Console.WriteLine("编辑限制已应用");
                Console.WriteLine("编辑限制演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"编辑限制演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 内容保护示例
        /// </summary>
        static void ContentProtectionDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 创建表单文档
                var range = document.Range();
                range.Text = "员工信息表\n\n姓名：________________\n部门：________________\n职位：________________\n薪资：________________";

                // 添加书签保护
                range.Collapse(WdCollapseDirection.wdCollapseStart);
                range.Text = "【受保护内容开始】";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 添加一些内容
                range.Text = "\n\n机密信息：这部分内容受到保护。";
                var confidentialRange = document.Range(range.Start - 20, range.End);

                // 为机密内容添加书签
                var bookmark = document.Bookmarks.Add("ConfidentialSection", confidentialRange);

                // 保护书签内容
                document.Bookmarks["ConfidentialSection"].Range.Editors.Add(WdEditorType.wdEditorOwners);

                // 添加表单字段保护
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n\n表单字段：";
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 添加受保护的表单字段
                var formField = range.FormFields.Add(range, WdFieldType.wdFieldFormTextInput);
                formField.Name = "ProtectedField";
                formField.TextInput.Default = "受保护的输入字段";
                formField.TextInput.EditType(WdTextFormFieldType.wdRegularText, "默认值", null, true); // 只读

                // 应用保护
                document.Protect(
                    protectionType: WdProtectionType.wdAllowOnlyFormFields, // 仅允许表单字段编辑
                    noReset: false,
                    password: "FormPass456");

                Console.WriteLine("内容保护已应用");
                Console.WriteLine("内容保护演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"内容保护演示出错: {ex.Message}");
            }
        }


        /// <summary>
        /// 文档权限管理示例
        /// </summary>
        static void DocumentPermissionManagementDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加文档内容
                document.Range().Text = "受限文档内容\n\n只有授权用户可以访问此文档。";

                // 设置文档权限（需要IRM - Information Rights Management支持）
                try
                {
                    // 检查是否支持权限管理
                    if (document.Permission.Enabled)
                    {
                        // 启用权限管理
                        document.Permission.Enabled = true;

                        // 添加用户权限
                        var userPermission = document.Permission.Add(
                            "user@example.com",
                           (int)MsoPermission.msoPermissionRead + (int)MsoPermission.msoPermissionEdit);

                        // 设置权限到期时间
                        userPermission.ExpirationDate = DateTime.Now.AddDays(30);

                        Console.WriteLine("文档权限已设置");
                    }
                    else
                    {
                        Console.WriteLine("当前系统不支持文档权限管理");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"设置文档权限时出错: {ex.Message}");
                }

                Console.WriteLine("文档权限管理演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文档权限管理演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 实际应用示例
        /// </summary>
        static void RealWorldApplicationDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                app.Visible = true;

                try
                {
                    var document = app.ActiveDocument;

                    // 设置文档属性
                    document.Title = "保密协议";
                    document.Subject = "商业机密保护";
                    document.Author = "法务部";
                    document.Company = "ABC有限公司";

                    // 创建合同标题
                    var titleRange = document.Range();
                    titleRange.Text = "保密协议\n";
                    titleRange.Font.Name = "微软雅黑";
                    titleRange.Font.Size = 18;
                    titleRange.Font.Bold = true;
                    titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    titleRange.ParagraphFormat.SpaceAfter = 24;

                    // 添加合同正文
                    var contentRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                    contentRange.Text = "本协议由以下双方于____年____月____日签署：\n\n";
                    contentRange.Font.Name = "宋体";
                    contentRange.Font.Size = 12;

                    // 甲方信息（可编辑）
                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "甲方（披露方）：\n";
                    contentRange.Font.Bold = true;

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "公司名称：________________________\n";
                    contentRange.Font.Bold = false;

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "地址：___________________________\n";

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "授权代表：_______________________\n";

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "职务：___________________________\n";

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "签字：___________________________\n";

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "日期：_______年____月____日\n\n";

                    // 乙方信息（可编辑）
                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "乙方（接收方）：\n";
                    contentRange.Font.Bold = true;

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "公司名称：________________________\n";
                    contentRange.Font.Bold = false;

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "地址：___________________________\n";

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "授权代表：_______________________\n";

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "职务：___________________________\n";

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "签字：___________________________\n";

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "日期：_______年____月____日\n\n";

                    // 合同条款（受保护）
                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "第一条 保密信息的定义\n";
                    contentRange.Font.Bold = true;

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "1.1 保密信息指甲方提供给乙方的任何技术、商业或其他信息...\n\n";
                    contentRange.Font.Bold = false;

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "第二条 保密义务\n";
                    contentRange.Font.Bold = true;

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "2.1 乙方应对保密信息严格保密...\n\n";
                    contentRange.Font.Bold = false;

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "第三条 保密期限\n";
                    contentRange.Font.Bold = true;

                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "3.1 本协议的保密期限为【    】年...\n\n";
                    contentRange.Font.Bold = false;

                    // 添加签名区域
                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "\n\n【以下无正文】\n";
                    contentRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                    // 添加签名行
                    contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    contentRange.Text = "\n\n甲方签字：______________    乙方签字：______________\n";
                    contentRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                    Console.WriteLine("安全合同文档创建完成");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"创建安全文档时出错: {ex.Message}");
                }

                Console.WriteLine("实际应用示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"实际应用示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 使用辅助类的完整示例
        /// </summary>
        static void CompleteExampleWithHelpers()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                app.Visible = false; // 隐藏Word窗口

                var document = app.ActiveDocument;

                // 创建文档保护助手
                var protectionHelper = new DocumentProtectionHelper(document);

                // 评估密码强度
                var passwordResult = protectionHelper.EvaluatePasswordStrength("MySecurePass123!");
                Console.WriteLine(passwordResult.GenerateReport());

                // 创建内容保护管理器
                var contentProtectionManager = new ContentProtectionManager(document);

                // 创建机密内容区域
                var allowedEditors = new List<string> { "manager@company.com", "admin@company.com" };
                contentProtectionManager.CreateConfidentialSection(
                    "这部分包含机密的财务信息，只有授权人员可以查看和编辑。",
                    allowedEditors);

                // 创建可编辑区域
                contentProtectionManager.CreateEditableSection(
                    "请在此处填写您的评论：\n______________________________________________\n______________________________________________",
                    WdEditorType.wdEditorEveryone);

                // 创建数字签名管理器
                var signatureManager = new DigitalSignatureManager(app, document);

                // 检查可用的签名提供商
                var providers = signatureManager.GetAvailableSignatureProviders();
                Console.WriteLine($"找到 {providers.Count} 个签名提供商");

                // 创建权限管理器
                var permissionManager = new PermissionManager(document);

                // 检查权限管理是否可用
                bool isPermissionAvailable = permissionManager.IsPermissionManagementAvailable();
                Console.WriteLine($"权限管理是否可用: {isPermissionAvailable}");

                // 创建安全文档构建器
                var secureDocumentBuilder = new SecureDocumentBuilder(app, document);

                // 验证安全设置
                var validationResult = secureDocumentBuilder.ValidateSecuritySettings();
                Console.WriteLine(validationResult.GenerateReport());

                // 生成安全报告
                var securityReport = secureDocumentBuilder.GenerateSecurityReport();
                Console.WriteLine("安全报告已生成");

                // 保存文档
                string filePath = Path.Combine(Path.GetTempPath(), "SecureDocumentWithHelpers.docx");

                // 设置保护设置
                var protectionSettings = new DocumentProtectionSettings
                {
                    CompanyName = "ABC有限公司",
                    OpenPassword = "OpenPass123",
                    ModifyPassword = "EditPass456",
                    ProtectionType = WdProtectionType.wdAllowOnlyReading,
                    ProtectionPassword = "ProtectPass789"
                };

                // 保存安全文档
                bool saved = secureDocumentBuilder.SaveSecureDocument(filePath, protectionSettings);
                if (saved)
                {
                    Console.WriteLine($"安全文档已保存: {filePath}");
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
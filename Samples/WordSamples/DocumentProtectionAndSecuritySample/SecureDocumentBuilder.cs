//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;

namespace DocumentProtectionAndSecuritySample
{
    /// <summary>
    /// 安全文档构建器类
    /// </summary>
    public class SecureDocumentBuilder
    {
        private readonly IWordApplication _application;
        private readonly IWordDocument _document;
        private readonly DocumentProtectionHelper _protectionHelper;
        private readonly ContentProtectionManager _contentProtectionManager;
        private readonly DigitalSignatureManager _signatureManager;
        private readonly PermissionManager _permissionManager;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="application">Word应用程序对象</param>
        /// <param name="document">Word文档对象</param>
        public SecureDocumentBuilder(IWordApplication? Application, IWordDocument document)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _document = document ?? throw new ArgumentNullException(nameof(document));

            _protectionHelper = new DocumentProtectionHelper(document);
            _contentProtectionManager = new ContentProtectionManager(document);
            _signatureManager = new DigitalSignatureManager(application, document);
            _permissionManager = new PermissionManager(document);
        }

        /// <summary>
        /// 创建安全合同文档
        /// </summary>
        /// <param name="title">合同标题</param>
        /// <param name="parties">签约方</param>
        /// <param name="terms">合同条款</param>
        /// <param name="protectionSettings">保护设置</param>
        /// <returns>是否创建成功</returns>
        public bool CreateSecureContract(
            string title,
            List<ContractParty> parties,
            List<ContractTerm> terms,
            DocumentProtectionSettings protectionSettings)
        {
            try
            {
                Console.WriteLine("开始创建安全合同文档...");

                // 设置文档属性
                _document.Title = title;
                _document.Subject = "商业合同";
                _document.Author = "法务部";
                _document.Company = protectionSettings.CompanyName;

                // 创建合同标题
                var titleRange = _document.Range();
                titleRange.Text = $"{title}\n";
                titleRange.Font.Name = "微软雅黑";
                titleRange.Font.Size = 18;
                titleRange.Font.Bold = true;
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                titleRange.ParagraphFormat.SpaceAfter = 24;

                // 添加合同签署方信息
                AddContractParties(parties);

                // 添加合同条款
                AddContractTerms(terms);

                // 添加签名区域
                AddSignatureArea(parties);

                // 应用内容保护
                if (protectionSettings.UseContentProtection)
                {
                    ApplyContentProtection(protectionSettings);
                }

                // 应用文档保护
                if (protectionSettings.UseDocumentProtection)
                {
                    ApplyDocumentProtection(protectionSettings);
                }

                // 应用权限管理
                if (protectionSettings.UsePermissionManagement)
                {
                    ApplyPermissionManagement(protectionSettings);
                }

                Console.WriteLine("安全合同文档创建完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建安全合同文档时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 添加合同签署方信息
        /// </summary>
        /// <param name="parties">签约方列表</param>
        private void AddContractParties(List<ContractParty> parties)
        {
            var contentRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
            contentRange.Text = $"本协议由以下双方于____年____月____日签署：\n\n";
            contentRange.Font.Name = "宋体";
            contentRange.Font.Size = 12;

            foreach (var party in parties)
            {
                contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                contentRange.Text = $"{party.PartyType}（{party.PartyRole}）：\n";
                contentRange.Font.Bold = true;

                contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                contentRange.Text = $"公司名称：{party.CompanyName}\n";
                contentRange.Font.Bold = false;

                contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                contentRange.Text = $"地址：{party.Address}\n";

                contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                contentRange.Text = $"授权代表：{party.AuthorizedRepresentative}\n";

                contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                contentRange.Text = $"职务：{party.Position}\n";

                contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                contentRange.Text = $"签字：___________________________\n";

                contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                contentRange.Text = $"日期：_______年____月____日\n\n";
            }
        }

        /// <summary>
        /// 添加合同条款
        /// </summary>
        /// <param name="terms">合同条款列表</param>
        private void AddContractTerms(List<ContractTerm> terms)
        {
            var contentRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);

            foreach (var term in terms)
            {
                contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                contentRange.Text = $"{term.TermNumber} {term.TermTitle}\n";
                contentRange.Font.Bold = true;

                contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                contentRange.Text = $"{term.TermContent}\n\n";
                contentRange.Font.Bold = false;
            }
        }

        /// <summary>
        /// 添加签名区域
        /// </summary>
        /// <param name="parties">签约方列表</param>
        private void AddSignatureArea(List<ContractParty> parties)
        {
            var contentRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);

            // 添加无正文标记
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "\n\n【以下无正文】\n";
            contentRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            // 添加签名行
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "\n\n";
            contentRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

            var signatureLines = parties.Select(p => new SignatureLineDefinition
            {
                Title = p.PartyType,
                SuggestedSigner = p.AuthorizedRepresentative,
                SuggestedSignerLine2 = p.Position,
                SuggestedSignerEmail = p.Email
            }).ToList();

            _signatureManager.AddMultipleSignatureLines(signatureLines);
        }

        /// <summary>
        /// 应用内容保护
        /// </summary>
        /// <param name="settings">保护设置</param>
        private void ApplyContentProtection(DocumentProtectionSettings settings)
        {
            try
            {
                // 创建机密内容区域
                if (settings.ConfidentialSections != null)
                {
                    foreach (var section in settings.ConfidentialSections)
                    {
                        _contentProtectionManager.CreateConfidentialSection(
                            section.Content,
                            section.AllowedEditors);
                    }
                }

                // 创建可编辑区域
                if (settings.EditableSections != null)
                {
                    foreach (var section in settings.EditableSections)
                    {
                        _contentProtectionManager.CreateEditableSection(
                            section.Content,
                            section.EditorType,
                            section.EditorName);
                    }
                }

                Console.WriteLine("内容保护已应用");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用内容保护时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 应用文档保护
        /// </summary>
        /// <param name="settings">保护设置</param>
        private void ApplyDocumentProtection(DocumentProtectionSettings settings)
        {
            try
            {
                // 设置密码
                if (!string.IsNullOrEmpty(settings.OpenPassword))
                {
                    _protectionHelper.SetOpenPassword(settings.OpenPassword);
                }

                if (!string.IsNullOrEmpty(settings.ModifyPassword))
                {
                    _protectionHelper.SetModifyPassword(settings.ModifyPassword);
                }

                // 应用保护
                _protectionHelper.ApplyProtection(
                    settings.ProtectionType,
                    settings.ProtectionPassword);

                Console.WriteLine("文档保护已应用");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用文档保护时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 应用权限管理
        /// </summary>
        /// <param name="settings">保护设置</param>
        private void ApplyPermissionManagement(DocumentProtectionSettings settings)
        {
            try
            {
                if (settings.AllowedUsers != null && settings.AllowedUsers.Any())
                {
                    _permissionManager.CreateStandardCorporatePolicy(
                        settings.AllowedUsers,
                        settings.PermissionExpirationDays);
                }

                Console.WriteLine("权限管理已应用");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用权限管理时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 创建受保护的表单文档
        /// </summary>
        /// <param name="title">表单标题</param>
        /// <param name="formFields">表单字段</param>
        /// <param name="protectionSettings">保护设置</param>
        /// <returns>是否创建成功</returns>
        public bool CreateProtectedForm(
            string title,
            List<FormFieldDefinition> formFields,
            DocumentProtectionSettings protectionSettings)
        {
            try
            {
                Console.WriteLine("开始创建受保护的表单文档...");

                // 设置文档属性
                _document.Title = title;
                _document.Subject = "表单文档";
                _document.Author = "表单系统";
                _document.Company = protectionSettings.CompanyName;

                // 创建受保护的表单文档
                _contentProtectionManager.CreateProtectedFormDocument(title, formFields);

                // 应用文档保护
                if (protectionSettings.UseDocumentProtection)
                {
                    ApplyDocumentProtection(protectionSettings);
                }

                // 应用权限管理
                if (protectionSettings.UsePermissionManagement)
                {
                    ApplyPermissionManagement(protectionSettings);
                }

                Console.WriteLine("受保护的表单文档创建完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建受保护的表单文档时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 验证文档安全设置
        /// </summary>
        /// <returns>验证结果</returns>
        public DocumentSecurityValidationResult ValidateSecuritySettings()
        {
            var result = new DocumentSecurityValidationResult();

            try
            {
                Console.WriteLine("开始验证文档安全设置...");

                // 检查文档保护状态
                result.ProtectionStatus = _protectionHelper.CheckProtectionStatus();

                // 检查权限管理状态
                result.PermissionStatus = _permissionManager.GetPermissionManagementStatus();

                // 验证签名
                result.SignatureValidation = _signatureManager.ValidateSignatures();

                // 获取受保护内容信息
                result.ProtectedContent = _contentProtectionManager.GetAllProtectedContentInfo();

                result.IsValid = true;
                Console.WriteLine("文档安全设置验证完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"验证文档安全设置时出错: {ex.Message}");
                result.ErrorMessage = ex.Message;
                result.IsValid = false;
            }

            return result;
        }

        /// <summary>
        /// 生成安全报告
        /// </summary>
        /// <returns>安全报告</returns>
        public DocumentSecurityReport GenerateSecurityReport()
        {
            var report = new DocumentSecurityReport
            {
                DocumentTitle = _document.Name,
                GeneratedDate = DateTime.Now
            };

            try
            {
                // 获取保护状态
                report.ProtectionStatus = _protectionHelper.CheckProtectionStatus();

                // 获取权限状态
                report.PermissionStatus = _permissionManager.GetPermissionManagementStatus();

                // 获取签名信息
                report.Signatures = _signatureManager.GetDocumentSignatures();

                // 获取受保护内容
                report.ProtectedContent = _contentProtectionManager.GetAllProtectedContentInfo();

                // 生成权限报告
                report.PermissionReport = _permissionManager.CreatePermissionReport();

                Console.WriteLine("安全报告生成完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成安全报告时出错: {ex.Message}");
                report.ErrorMessage = ex.Message;
            }

            return report;
        }

        /// <summary>
        /// 保存安全文档
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="protectionSettings">保护设置</param>
        /// <returns>是否保存成功</returns>
        public bool SaveSecureDocument(string filePath, DocumentProtectionSettings protectionSettings)
        {
            try
            {
                // 保存文档
                _document.SaveAs(
                    fileName: filePath,
                    password: protectionSettings.OpenPassword,
                    writePassword: protectionSettings.ModifyPassword
                );

                Console.WriteLine($"安全文档已保存: {filePath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存安全文档时出错: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// 合同签署方类
    /// </summary>
    public class ContractParty
    {
        /// <summary>
        /// 签署方类型（甲方、乙方等）
        /// </summary>
        public string PartyType { get; set; }

        /// <summary>
        /// 签署方角色
        /// </summary>
        public string PartyRole { get; set; }

        /// <summary>
        /// 公司名称
        /// </summary>
        public string CompanyName { get; set; }

        /// <summary>
        /// 地址
        /// </summary>
        public string Address { get; set; }

        /// <summary>
        /// 授权代表
        /// </summary>
        public string AuthorizedRepresentative { get; set; }

        /// <summary>
        /// 职务
        /// </summary>
        public string Position { get; set; }

        /// <summary>
        /// 邮箱
        /// </summary>
        public string Email { get; set; }
    }

    /// <summary>
    /// 合同条款类
    /// </summary>
    public class ContractTerm
    {
        /// <summary>
        /// 条款编号
        /// </summary>
        public string TermNumber { get; set; }

        /// <summary>
        /// 条款标题
        /// </summary>
        public string TermTitle { get; set; }

        /// <summary>
        /// 条款内容
        /// </summary>
        public string TermContent { get; set; }
    }

    /// <summary>
    /// 文档保护设置类
    /// </summary>
    public class DocumentProtectionSettings
    {
        /// <summary>
        /// 公司名称
        /// </summary>
        public string CompanyName { get; set; } = "ABC有限公司";

        /// <summary>
        /// 是否使用内容保护
        /// </summary>
        public bool UseContentProtection { get; set; } = true;

        /// <summary>
        /// 是否使用文档保护
        /// </summary>
        public bool UseDocumentProtection { get; set; } = true;

        /// <summary>
        /// 是否使用权限管理
        /// </summary>
        public bool UsePermissionManagement { get; set; } = false;

        /// <summary>
        /// 打开密码
        /// </summary>
        public string OpenPassword { get; set; }

        /// <summary>
        /// 修改密码
        /// </summary>
        public string ModifyPassword { get; set; }

        /// <summary>
        /// 保护类型
        /// </summary>
        public WdProtectionType ProtectionType { get; set; } = WdProtectionType.wdAllowOnlyReading;

        /// <summary>
        /// 保护密码
        /// </summary>
        public string ProtectionPassword { get; set; }

        /// <summary>
        /// 机密内容区域
        /// </summary>
        public List<ConfidentialSection> ConfidentialSections { get; set; }

        /// <summary>
        /// 可编辑区域
        /// </summary>
        public List<EditableSection> EditableSections { get; set; }

        /// <summary>
        /// 允许的用户列表
        /// </summary>
        public List<string> AllowedUsers { get; set; }

        /// <summary>
        /// 权限过期天数
        /// </summary>
        public int PermissionExpirationDays { get; set; } = 30;
    }

    /// <summary>
    /// 机密内容区域类
    /// </summary>
    public class ConfidentialSection
    {
        /// <summary>
        /// 内容
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// 允许的编辑者列表
        /// </summary>
        public List<WdEditorType> AllowedEditors { get; set; }
    }

    /// <summary>
    /// 可编辑区域类
    /// </summary>
    public class EditableSection
    {
        /// <summary>
        /// 内容
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// 编辑者类型
        /// </summary>
        public WdEditorType EditorType { get; set; } = WdEditorType.wdEditorEveryone;

        /// <summary>
        /// 编辑者名称
        /// </summary>
        public string EditorName { get; set; }
    }

    /// <summary>
    /// 文档安全验证结果类
    /// </summary>
    public class DocumentSecurityValidationResult
    {
        /// <summary>
        /// 是否有效
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 保护状态
        /// </summary>
        public DocumentProtectionStatus ProtectionStatus { get; set; }

        /// <summary>
        /// 权限状态
        /// </summary>
        public PermissionManagementStatus PermissionStatus { get; set; }

        /// <summary>
        /// 签名验证结果
        /// </summary>
        public SignatureValidationResult SignatureValidation { get; set; }

        /// <summary>
        /// 受保护内容
        /// </summary>
        public List<ProtectedContentInfo> ProtectedContent { get; set; } = new List<ProtectedContentInfo>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成验证结果报告
        /// </summary>
        /// <returns>验证结果报告</returns>
        public string GenerateReport()
        {
            if (!IsValid)
            {
                return $"文档安全验证失败: {ErrorMessage}";
            }

            return $"文档安全验证结果报告:\n" +
                   $"  验证状态: 通过\n" +
                   $"  {ProtectionStatus.GenerateReport()}\n" +
                   $"  {PermissionStatus.GenerateReport()}\n" +
                   $"  {SignatureValidation.GenerateReport()}\n" +
                   $"  受保护内容项数: {ProtectedContent.Count}";
        }
    }

    /// <summary>
    /// 文档安全报告类
    /// </summary>
    public class DocumentSecurityReport
    {
        /// <summary>
        /// 文档标题
        /// </summary>
        public string DocumentTitle { get; set; }

        /// <summary>
        /// 生成日期
        /// </summary>
        public DateTime GeneratedDate { get; set; }

        /// <summary>
        /// 保护状态
        /// </summary>
        public DocumentProtectionStatus ProtectionStatus { get; set; }

        /// <summary>
        /// 权限状态
        /// </summary>
        public PermissionManagementStatus PermissionStatus { get; set; }

        /// <summary>
        /// 签名列表
        /// </summary>
        public List<SignatureInfo> Signatures { get; set; } = new List<SignatureInfo>();

        /// <summary>
        /// 受保护内容
        /// </summary>
        public List<ProtectedContentInfo> ProtectedContent { get; set; } = new List<ProtectedContentInfo>();

        /// <summary>
        /// 权限报告
        /// </summary>
        public PermissionReport PermissionReport { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成安全报告
        /// </summary>
        /// <returns>安全报告</returns>
        public string GenerateReport()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"生成安全报告失败: {ErrorMessage}";
            }

            var signatureReports = Signatures.Select(s => s.GenerateReport()).ToList();
            var signaturesReport = signatureReports.Any() ? string.Join("\n\n", signatureReports) : "无签名信息";

            var contentReports = ProtectedContent.Select(c => c.GenerateReport()).ToList();
            var contentReport = contentReports.Any() ? string.Join("\n\n", contentReports) : "无受保护内容";

            return $"文档安全报告\n" +
                   $"文档标题: {DocumentTitle}\n" +
                   $"生成日期: {GeneratedDate:yyyy-MM-dd HH:mm:ss}\n\n" +
                   $"{ProtectionStatus.GenerateReport()}\n\n" +
                   $"{PermissionStatus.GenerateReport()}\n\n" +
                   $"签名信息:\n{signaturesReport}\n\n" +
                   $"受保护内容:\n{contentReport}\n\n" +
                   $"{PermissionReport.GenerateReport()}";
        }
    }
}
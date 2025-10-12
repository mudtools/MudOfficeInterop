using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentProtectionAndSecuritySample
{
    /// <summary>
    /// 数字签名管理器类
    /// </summary>
    public class DigitalSignatureManager
    {
        private readonly IWordApplication _application;
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="application">Word应用程序对象</param>
        /// <param name="document">Word文档对象</param>
        public DigitalSignatureManager(IWordApplication application, IWordDocument document)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 检查可用的签名提供商
        /// </summary>
        /// <returns>签名提供商信息列表</returns>
        public List<SignatureProviderInfo> GetAvailableSignatureProviders()
        {
            var providers = new List<SignatureProviderInfo>();

            try
            {
                var signatureProviders = _application.SignatureProviders;
                Console.WriteLine($"可用签名提供商数量: {signatureProviders.Count}");

                for (int i = 1; i <= signatureProviders.Count; i++)
                {
                    var provider = signatureProviders.Item(i);
                    var providerInfo = new SignatureProviderInfo
                    {
                        Id = provider.Id,
                        Name = provider.Name,
                        Version = provider.Version
                    };
                    providers.Add(providerInfo);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取签名提供商信息时出错: {ex.Message}");
            }

            return providers;
        }

        /// <summary>
        /// 添加签名行
        /// </summary>
        /// <param name="position">签名行位置</param>
        /// <param name="suggestedSigner">建议签名者</param>
        /// <param name="suggestedSignerLine2">建议签名者第二行</param>
        /// <param name="suggestedSignerEmail">建议签名者邮箱</param>
        /// <returns>签名行对象</returns>
        public IWordSignatureLine AddSignatureLine(
            IWordRange position,
            string suggestedSigner,
            string suggestedSignerLine2 = "",
            string suggestedSignerEmail = "")
        {
            try
            {
                // 添加签名行
                var shape = _document.Shapes.AddSignatureLine(
                    new object(), // SignatureLineSpec
                    position.Left, position.Top, 200, 50);

                var signatureLine = shape.SignatureLine;
                signatureLine.SuggestedSigner = suggestedSigner;
                signatureLine.SuggestedSignerLine2 = suggestedSignerLine2;
                signatureLine.SuggestedSignerEmail = suggestedSignerEmail;

                Console.WriteLine($"签名行已添加: {suggestedSigner}");
                return signatureLine;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加签名行时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 添加多个签名行
        /// </summary>
        /// <param name="signatureLines">签名行定义列表</param>
        /// <returns>是否添加成功</returns>
        public bool AddMultipleSignatureLines(List<SignatureLineDefinition> signatureLines)
        {
            try
            {
                foreach (var signatureLine in signatureLines)
                {
                    var range = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                    range.Collapse(WdCollapseDirection.wdCollapseEnd);
                    range.Text = $"\n{signatureLine.Title}:\n";
                    range.Collapse(WdCollapseDirection.wdCollapseEnd);

                    AddSignatureLine(
                        range,
                        signatureLine.SuggestedSigner,
                        signatureLine.SuggestedSignerLine2,
                        signatureLine.SuggestedSignerEmail);
                }

                Console.WriteLine($"已添加 {signatureLines.Count} 个签名行");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加多个签名行时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取文档签名信息
        /// </summary>
        /// <returns>签名信息列表</returns>
        public List<SignatureInfo> GetDocumentSignatures()
        {
            var signatures = new List<SignatureInfo>();

            try
            {
                var documentSignatures = _document.Signatures;
                Console.WriteLine($"文档签名数量: {documentSignatures.Count}");

                for (int i = 1; i <= documentSignatures.Count; i++)
                {
                    var signature = documentSignatures.Item(i);
                    var signatureInfo = new SignatureInfo
                    {
                        Id = signature.Id,
                        Signer = signature.Signer,
                        SignerEmail = signature.SignerEmail,
                        SignatureDate = signature.SignatureDate,
                        IsValid = signature.IsValid,
                        IsSigned = signature.IsSigned
                    };
                    signatures.Add(signatureInfo);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取文档签名信息时出错: {ex.Message}");
            }

            return signatures;
        }

        /// <summary>
        /// 验证文档签名
        /// </summary>
        /// <returns>验证结果</returns>
        public SignatureValidationResult ValidateSignatures()
        {
            var result = new SignatureValidationResult();

            try
            {
                var signatures = _document.Signatures;
                result.TotalSignatures = signatures.Count;
                result.ValidSignatures = 0;
                result.InvalidSignatures = 0;

                for (int i = 1; i <= signatures.Count; i++)
                {
                    var signature = signatures.Item(i);
                    if (signature.IsValid)
                    {
                        result.ValidSignatures++;
                    }
                    else
                    {
                        result.InvalidSignatures++;
                    }
                }

                result.IsDocumentValid = result.InvalidSignatures == 0 && result.TotalSignatures > 0;
                Console.WriteLine($"文档签名验证完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"验证文档签名时出错: {ex.Message}");
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 创建待签名文档
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>是否创建成功</returns>
        public bool CreateDocumentForSigning(string filePath)
        {
            try
            {
                // 保存文档以准备签名
                _document.SaveAs2(filePath);
                Console.WriteLine($"待签名文档已保存: {filePath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建待签名文档时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 添加时间戳
        /// </summary>
        /// <param name="range">文档范围</param>
        /// <returns>是否添加成功</returns>
        public bool AddTimestamp(IWordRange range)
        {
            try
            {
                range.Fields.Add(range, WdFieldType.wdFieldDate);
                Console.WriteLine("时间戳已添加");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加时间戳时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建合同签名页
        /// </summary>
        /// <param name="parties">签约方列表</param>
        /// <returns>是否创建成功</returns>
        public bool CreateContractSignaturePage(List<SigningParty> parties)
        {
            try
            {
                var range = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                
                // 添加签名页标题
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n\n签署页\n";
                range.Font.Name = "微软雅黑";
                range.Font.Size = 14;
                range.Font.Bold = 1;
                range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                
                // 为每个签约方添加签名行
                foreach (var party in parties)
                {
                    range.Collapse(WdCollapseDirection.wdCollapseEnd);
                    range.Text = $"\n{party.PartyName}:\n";
                    range.Font.Bold = 1;
                    range.Collapse(WdCollapseDirection.wdCollapseEnd);
                    
                    range.Text = $"签名：____________________    日期：____年____月____日\n";
                    range.Font.Bold = 0;
                    range.Collapse(WdCollapseDirection.wdCollapseEnd);
                }
                
                // 添加无正文标记
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n\n【以下无正文】\n";
                range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                
                Console.WriteLine("合同签名页已创建");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建合同签名页时出错: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// 签名提供商信息类
    /// </summary>
    public class SignatureProviderInfo
    {
        /// <summary>
        /// ID
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 版本
        /// </summary>
        public string Version { get; set; }

        /// <summary>
        /// 生成提供商信息报告
        /// </summary>
        /// <returns>信息报告</returns>
        public string GenerateReport()
        {
            return $"签名提供商信息:\n" +
                   $"  ID: {Id}\n" +
                   $"  名称: {Name}\n" +
                   $"  版本: {Version}";
        }
    }

    /// <summary>
    /// 签名行定义类
    /// </summary>
    public class SignatureLineDefinition
    {
        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 建议签名者
        /// </summary>
        public string SuggestedSigner { get; set; }

        /// <summary>
        /// 建议签名者第二行
        /// </summary>
        public string SuggestedSignerLine2 { get; set; }

        /// <summary>
        /// 建议签名者邮箱
        /// </summary>
        public string SuggestedSignerEmail { get; set; }
    }

    /// <summary>
    /// 签名信息类
    /// </summary>
    public class SignatureInfo
    {
        /// <summary>
        /// ID
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 签名者
        /// </summary>
        public string Signer { get; set; }

        /// <summary>
        /// 签名者邮箱
        /// </summary>
        public string SignerEmail { get; set; }

        /// <summary>
        /// 签名日期
        /// </summary>
        public DateTime SignatureDate { get; set; }

        /// <summary>
        /// 是否有效
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 是否已签名
        /// </summary>
        public bool IsSigned { get; set; }

        /// <summary>
        /// 生成签名信息报告
        /// </summary>
        /// <returns>信息报告</returns>
        public string GenerateReport()
        {
            return $"签名信息:\n" +
                   $"  签名者: {Signer}\n" +
                   $"  邮箱: {SignerEmail}\n" +
                   $"  签名日期: {SignatureDate:yyyy-MM-dd HH:mm:ss}\n" +
                   $"  是否有效: {IsValid}\n" +
                   $"  是否已签名: {IsSigned}";
        }
    }

    /// <summary>
    /// 签名验证结果类
    /// </summary>
    public class SignatureValidationResult
    {
        /// <summary>
        /// 总签名数
        /// </summary>
        public int TotalSignatures { get; set; }

        /// <summary>
        /// 有效签名数
        /// </summary>
        public int ValidSignatures { get; set; }

        /// <summary>
        /// 无效签名数
        /// </summary>
        public int InvalidSignatures { get; set; }

        /// <summary>
        /// 文档是否有效
        /// </summary>
        public bool IsDocumentValid { get; set; }

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
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"签名验证失败: {ErrorMessage}";
            }

            return $"签名验证结果报告:\n" +
                   $"  总签名数: {TotalSignatures}\n" +
                   $"  有效签名数: {ValidSignatures}\n" +
                   $"  无效签名数: {InvalidSignatures}\n" +
                   $"  文档是否有效: {IsDocumentValid}";
        }
    }

    /// <summary>
    /// 签约方类
    /// </summary>
    public class SigningParty
    {
        /// <summary>
        /// 签约方名称
        /// </summary>
        public string PartyName { get; set; }

        /// <summary>
        /// 授权代表
        /// </summary>
        public string AuthorizedRepresentative { get; set; }

        /// <summary>
        /// 职务
        /// </summary>
        public string Position { get; set; }
    }
}
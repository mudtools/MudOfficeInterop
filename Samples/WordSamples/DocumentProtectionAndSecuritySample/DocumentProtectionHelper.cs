using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentProtectionAndSecuritySample
{
    /// <summary>
    /// 文档保护助手类
    /// </summary>
    public class DocumentProtectionHelper
    {
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public DocumentProtectionHelper(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 设置文档打开密码
        /// </summary>
        /// <param name="password">密码</param>
        /// <returns>是否设置成功</returns>
        public bool SetOpenPassword(string password)
        {
            try
            {
                _document.Password = password;
                Console.WriteLine("文档打开密码已设置");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置文档打开密码时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 设置文档修改密码
        /// </summary>
        /// <param name="password">密码</param>
        /// <returns>是否设置成功</returns>
        public bool SetModifyPassword(string password)
        {
            try
            {
                _document.WritePassword = password;
                Console.WriteLine("文档修改密码已设置");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置文档修改密码时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 设置加密提供程序
        /// </summary>
        /// <param name="provider">加密提供程序名称</param>
        /// <returns>是否设置成功</returns>
        public bool SetEncryptionProvider(string provider)
        {
            try
            {
                _document.EncryptionProvider = provider;
                Console.WriteLine($"加密提供程序已设置为: {provider}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置加密提供程序时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 应用文档保护
        /// </summary>
        /// <param name="protectionType">保护类型</param>
        /// <param name="password">保护密码</param>
        /// <param name="noReset">是否不重置现有保护</param>
        /// <returns>是否应用成功</returns>
        public bool ApplyProtection(WdProtectionType protectionType, string password = null, bool noReset = false)
        {
            try
            {
                _document.Protect(
                    Type: protectionType,
                    NoReset: noReset,
                    Password: password
                );

                Console.WriteLine($"文档保护已应用: {protectionType}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用文档保护时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 移除文档保护
        /// </summary>
        /// <param name="password">保护密码</param>
        /// <returns>是否移除成功</returns>
        public bool UnprotectDocument(string password = null)
        {
            try
            {
                _document.Unprotect(Password: password);
                Console.WriteLine("文档保护已移除");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"移除文档保护时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 检查文档是否受保护
        /// </summary>
        /// <returns>保护状态信息</returns>
        public DocumentProtectionStatus CheckProtectionStatus()
        {
            var status = new DocumentProtectionStatus();

            try
            {
                status.IsProtected = _document.ProtectionType != WdProtectionType.wdNoProtection;
                status.ProtectionType = _document.ProtectionType;
                status.HasOpenPassword = !string.IsNullOrEmpty(_document.Password);
                status.HasModifyPassword = !string.IsNullOrEmpty(_document.WritePassword);
                status.EditableRangesCount = _document.EditableRanges.Count;
                status.BookmarksCount = _document.Bookmarks.Count;
                status.SignaturesCount = _document.Signatures.Count;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"检查文档保护状态时出错: {ex.Message}");
                status.ErrorMessage = ex.Message;
            }

            return status;
        }

        /// <summary>
        /// 添加可编辑区域
        /// </summary>
        /// <param name="range">文档范围</param>
        /// <param name="editorType">编辑者类型</param>
        /// <param name="editorName">编辑者名称（可选）</param>
        /// <returns>是否添加成功</returns>
        public bool AddEditableRange(IWordRange range, WdEditorType editorType, string editorName = null)
        {
            try
            {
                var editableRange = _document.EditableRanges.Add(range);
                
                if (editorType == WdEditorType.wdEditorEveryone)
                {
                    editableRange.Editors.Add(WdEditorType.wdEditorEveryone);
                }
                else if (!string.IsNullOrEmpty(editorName))
                {
                    editableRange.Editors.Add(editorName);
                }

                Console.WriteLine("可编辑区域已添加");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加可编辑区域时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取文档保护类型描述
        /// </summary>
        /// <param name="protectionType">保护类型</param>
        /// <returns>保护类型描述</returns>
        public string GetProtectionTypeDescription(WdProtectionType protectionType)
        {
            switch (protectionType)
            {
                case WdProtectionType.wdNoProtection:
                    return "无保护";
                case WdProtectionType.wdAllowOnlyRevisions:
                    return "仅允许修订";
                case WdProtectionType.wdAllowOnlyComments:
                    return "仅允许批注";
                case WdProtectionType.wdAllowOnlyFormFields:
                    return "仅允许表单字段";
                case WdProtectionType.wdAllowOnlyReading:
                    return "仅允许阅读";
                default:
                    return "未知保护类型";
            }
        }

        /// <summary>
        /// 验证密码强度
        /// </summary>
        /// <param name="password">密码</param>
        /// <returns>密码强度评估</returns>
        public PasswordStrengthResult EvaluatePasswordStrength(string password)
        {
            var result = new PasswordStrengthResult
            {
                Password = password
            };

            if (string.IsNullOrEmpty(password))
            {
                result.Strength = "空密码";
                result.Score = 0;
                result.Recommendations = new List<string> { "请设置密码" };
                return result;
            }

            var recommendations = new List<string>();

            // 检查密码长度
            if (password.Length < 8)
            {
                result.Score -= 2;
                recommendations.Add("密码长度应至少8位");
            }
            else if (password.Length >= 12)
            {
                result.Score += 2;
            }

            // 检查是否包含数字
            if (!password.Any(char.IsDigit))
            {
                result.Score -= 1;
                recommendations.Add("密码应包含数字");
            }
            else
            {
                result.Score += 1;
            }

            // 检查是否包含大写字母
            if (!password.Any(char.IsUpper))
            {
                result.Score -= 1;
                recommendations.Add("密码应包含大写字母");
            }
            else
            {
                result.Score += 1;
            }

            // 检查是否包含小写字母
            if (!password.Any(char.IsLower))
            {
                result.Score -= 1;
                recommendations.Add("密码应包含小写字母");
            }
            else
            {
                result.Score += 1;
            }

            // 检查是否包含特殊字符
            if (!password.Any(c => !char.IsLetterOrDigit(c)))
            {
                result.Score -= 1;
                recommendations.Add("密码应包含特殊字符");
            }
            else
            {
                result.Score += 1;
            }

            // 评估强度
            if (result.Score <= 0)
            {
                result.Strength = "弱";
            }
            else if (result.Score <= 2)
            {
                result.Strength = "中等";
            }
            else if (result.Score <= 4)
            {
                result.Strength = "强";
            }
            else
            {
                result.Strength = "很强";
            }

            result.Recommendations = recommendations;
            return result;
        }
    }

    /// <summary>
    /// 文档保护状态信息类
    /// </summary>
    public class DocumentProtectionStatus
    {
        /// <summary>
        /// 是否受保护
        /// </summary>
        public bool IsProtected { get; set; }

        /// <summary>
        /// 保护类型
        /// </summary>
        public WdProtectionType ProtectionType { get; set; }

        /// <summary>
        /// 是否有打开密码
        /// </summary>
        public bool HasOpenPassword { get; set; }

        /// <summary>
        /// 是否有修改密码
        /// </summary>
        public bool HasModifyPassword { get; set; }

        /// <summary>
        /// 可编辑区域数量
        /// </summary>
        public int EditableRangesCount { get; set; }

        /// <summary>
        /// 书签数量
        /// </summary>
        public int BookmarksCount { get; set; }

        /// <summary>
        /// 数字签名数量
        /// </summary>
        public int SignaturesCount { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成保护状态报告
        /// </summary>
        /// <returns>状态报告</returns>
        public string GenerateReport()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"检查保护状态失败: {ErrorMessage}";
            }

            return $"文档保护状态报告:\n" +
                   $"  是否受保护: {IsProtected}\n" +
                   $"  保护类型: {ProtectionType}\n" +
                   $"  是否有打开密码: {HasOpenPassword}\n" +
                   $"  是否有修改密码: {HasModifyPassword}\n" +
                   $"  可编辑区域数量: {EditableRangesCount}\n" +
                   $"  书签数量: {BookmarksCount}\n" +
                   $"  数字签名数量: {SignaturesCount}";
        }
    }

    /// <summary>
    /// 密码强度评估结果类
    /// </summary>
    public class PasswordStrengthResult
    {
        /// <summary>
        /// 密码
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// 强度评分
        /// </summary>
        public int Score { get; set; } = 0;

        /// <summary>
        /// 强度描述
        /// </summary>
        public string Strength { get; set; }

        /// <summary>
        /// 建议列表
        /// </summary>
        public List<string> Recommendations { get; set; } = new List<string>();

        /// <summary>
        /// 生成密码强度报告
        /// </summary>
        /// <returns>强度报告</returns>
        public string GenerateReport()
        {
            var recommendations = Recommendations.Any() ? 
                string.Join("\n  - ", Recommendations) : 
                "密码强度良好";

            return $"密码强度评估报告:\n" +
                   $"  密码: {Password}\n" +
                   $"  强度评分: {Score}\n" +
                   $"  强度等级: {Strength}\n" +
                   $"  改进建议: \n  - {recommendations}";
        }
    }
}
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;
using System.Globalization;
using System.Text;

namespace FAQSample
{
    /// <summary>
    /// 兼容性帮助类
    /// </summary>
    public class CompatibilityHelper
    {
        /// <summary>
        /// Office版本信息
        /// </summary>
        public class OfficeVersionInfo
        {
            /// <summary>
            /// 版本号
            /// </summary>
            public string Version { get; set; }

            /// <summary>
            /// 版本名称
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// 发布年份
            /// </summary>
            public int ReleaseYear { get; set; }

            /// <summary>
            /// 是否支持
            /// </summary>
            public bool IsSupported { get; set; }
        }

        /// <summary>
        /// 获取支持的Office版本信息
        /// </summary>
        /// <returns>Office版本信息列表</returns>
        public static List<OfficeVersionInfo> GetSupportedOfficeVersions()
        {
            return new List<OfficeVersionInfo>
            {
                new OfficeVersionInfo
                {
                    Version = "16.0",
                    Name = "Office 2016/2019/365",
                    ReleaseYear = 2015,
                    IsSupported = true
                },
                new OfficeVersionInfo
                {
                    Version = "15.0",
                    Name = "Office 2013",
                    ReleaseYear = 2012,
                    IsSupported = true
                },
                new OfficeVersionInfo
                {
                    Version = "14.0",
                    Name = "Office 2010",
                    ReleaseYear = 2009,
                    IsSupported = true
                },
                new OfficeVersionInfo
                {
                    Version = "12.0",
                    Name = "Office 2007",
                    ReleaseYear = 2006,
                    IsSupported = false
                }
            };
        }

        /// <summary>
        /// 检查Office版本兼容性
        /// </summary>
        /// <returns>兼容性检查结果</returns>
        public static CompatibilityCheckResult CheckOfficeVersionCompatibility()
        {
            var result = new CompatibilityCheckResult();

            try
            {
                using var app = WordFactory.BlankWorkbook();
                var version = app.Version;
                result.DetectedVersion = version;

                var supportedVersions = GetSupportedOfficeVersions();
                var matchedVersion = supportedVersions.FirstOrDefault(v => v.Version == version);

                if (matchedVersion != null)
                {
                    result.IsCompatible = matchedVersion.IsSupported;
                    result.VersionInfo = matchedVersion;
                    result.Message = matchedVersion.IsSupported
                        ? $"检测到兼容版本: {matchedVersion.Name}"
                        : $"检测到不支持版本: {matchedVersion.Name}";
                }
                else
                {
                    result.IsCompatible = false;
                    result.Message = $"检测到未知版本: {version}";
                }
            }
            catch (Exception ex)
            {
                result.IsCompatible = false;
                result.Message = $"检查Office版本时出错: {ex.Message}";
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 多语言支持帮助类
        /// </summary>
        public class MultiLanguageSupportHelper
        {
            /// <summary>
            /// 应用程序区域设置
            /// </summary>
            public CultureInfo ApplicationCulture { get; set; } = CultureInfo.CurrentUICulture;

            /// <summary>
            /// 设置文档语言
            /// </summary>
            /// <param name="document">Word文档</param>
            /// <param name="language">语言</param>
            public void SetDocumentLanguage(IWordDocument document, CultureInfo language)
            {
                try
                {
                    var range = document.Range();
                    range.LanguageID = GetWdLanguageID(language);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"设置文档语言时出错: {ex.Message}");
                }
            }

            /// <summary>
            /// 获取Word语言ID
            /// </summary>
            /// <param name="culture">区域性信息</param>
            /// <returns>Word语言ID</returns>
            private WdLanguageID GetWdLanguageID(CultureInfo culture)
            {
                return culture.TwoLetterISOLanguageName.ToLower() switch
                {
                    "zh" => WdLanguageID.wdChinesePRC,
                    "en" => WdLanguageID.wdEnglishUS,
                    "ja" => WdLanguageID.wdJapanese,
                    "ko" => WdLanguageID.wdKorean,
                    "fr" => WdLanguageID.wdFrench,
                    "de" => WdLanguageID.wdGerman,
                    "es" => WdLanguageID.wdSpanish,
                    _ => WdLanguageID.wdEnglishUS
                };
            }

            /// <summary>
            /// 使用程序化方式而不是UI操作
            /// </summary>
            /// <param name="selection">选择区域</param>
            public void ApplyFormattingProgrammatically(IWordSelection selection)
            {
                // 正确：使用枚举常量
                selection.Font.Bold = 1;
                selection.Font.Italic = 1;
                selection.Font.Underline = (int)WdUnderline.wdUnderlineSingle;
            }

            /// <summary>
            /// 避免依赖特定语言的菜单项
            /// </summary>
            /// <param name="application">Word应用程序</param>
            public void AvoidLanguageSpecificCommands(IWordApplication application)
            {
                // 正确：使用程序化方式
                application.ScreenUpdating = false;

                // 避免：使用特定语言的命令
                // application.CommandBars.Execute("Save"); // 可能在不同语言版本中失败
            }
        }

        /// <summary>
        /// 测试不同Office版本的功能
        /// </summary>
        /// <returns>功能测试结果</returns>
        public static FeatureTestResult TestOfficeFeatures()
        {
            var result = new FeatureTestResult();

            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 测试基本功能
                result.BasicFeatures = TestBasicFeatures(document);

                // 测试高级功能
                result.AdvancedFeatures = TestAdvancedFeatures(document);

                result.Success = true;
                result.Message = "功能测试完成";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"功能测试时出错: {ex.Message}";
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 测试基本功能
        /// </summary>
        /// <param name="document">Word文档</param>
        /// <returns>基本功能测试结果</returns>
        private static Dictionary<string, bool> TestBasicFeatures(IWordDocument document)
        {
            var features = new Dictionary<string, bool>();

            try
            {
                // 测试文本操作
                document.Range().Text = "测试文本";
                features["文本操作"] = true;
            }
            catch
            {
                features["文本操作"] = false;
            }

            try
            {
                // 测试段落操作
                document.Paragraphs[1].Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                features["段落操作"] = true;
            }
            catch
            {
                features["段落操作"] = false;
            }

            try
            {
                // 测试字体操作
                document.Range().Font.Bold = 1;
                features["字体操作"] = true;
            }
            catch
            {
                features["字体操作"] = false;
            }

            return features;
        }

        /// <summary>
        /// 测试高级功能
        /// </summary>
        /// <param name="document">Word文档</param>
        /// <returns>高级功能测试结果</returns>
        private static Dictionary<string, bool> TestAdvancedFeatures(IWordDocument document)
        {
            var features = new Dictionary<string, bool>();

            try
            {
                // 测试表格操作
                var table = document.Tables.Add(document.Range(), 2, 2);
                features["表格操作"] = table != null;
            }
            catch
            {
                features["表格操作"] = false;
            }

            try
            {
                // 测试样式操作
                var style = document.Styles["标题 1"];
                features["样式操作"] = style != null;
            }
            catch
            {
                features["样式操作"] = false;
            }

            return features;
        }

        /// <summary>
        /// 生成兼容性报告
        /// </summary>
        /// <returns>兼容性报告</returns>
        public static string GenerateCompatibilityReport()
        {
            var report = new StringBuilder();
            report.AppendLine("=== 兼容性报告 ===");
            report.AppendLine();

            // Office版本兼容性
            report.AppendLine("1. 支持的Office版本:");
            var versions = GetSupportedOfficeVersions();
            foreach (var version in versions)
            {
                report.AppendLine($"   - {version.Name} ({version.Version}) - {(version.IsSupported ? "支持" : "不支持")}");
            }
            report.AppendLine();

            // 多语言支持
            report.AppendLine("2. 多语言支持建议:");
            report.AppendLine("   - 使用程序化方式而不是UI操作");
            report.AppendLine("   - 避免依赖特定语言的菜单项或对话框");
            report.AppendLine("   - 使用常量而不是硬编码的字符串");
            report.AppendLine();

            // 最佳实践
            report.AppendLine("3. 兼容性最佳实践:");
            report.AppendLine("   - 使用与目标环境相同或相近版本的Office进行测试");
            report.AppendLine("   - 注意新版本Office可能添加的API");
            report.AppendLine("   - 避免使用已弃用的功能");
            report.AppendLine();

            report.AppendLine("=================");

            return report.ToString();
        }
    }

    /// <summary>
    /// 兼容性检查结果类
    /// </summary>
    public class CompatibilityCheckResult
    {
        /// <summary>
        /// 是否兼容
        /// </summary>
        public bool IsCompatible { get; set; }

        /// <summary>
        /// 检测到的版本
        /// </summary>
        public string DetectedVersion { get; set; }

        /// <summary>
        /// 版本信息
        /// </summary>
        public CompatibilityHelper.OfficeVersionInfo VersionInfo { get; set; }

        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 功能测试结果类
    /// </summary>
    public class FeatureTestResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 基本功能测试结果
        /// </summary>
        public Dictionary<string, bool> BasicFeatures { get; set; } = new Dictionary<string, bool>();

        /// <summary>
        /// 高级功能测试结果
        /// </summary>
        public Dictionary<string, bool> AdvancedFeatures { get; set; } = new Dictionary<string, bool>();

        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }
}
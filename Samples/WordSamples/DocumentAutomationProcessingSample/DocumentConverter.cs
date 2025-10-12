using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentAutomationProcessingSample
{
    /// <summary>
    /// 文档转换器类
    /// </summary>
    public class DocumentConverter
    {
        /// <summary>
        /// 转换目录中的文档
        /// </summary>
        /// <param name="inputDirectory">输入目录</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="targetFormat">目标格式</param>
        /// <param name="filePattern">文件匹配模式</param>
        /// <returns>转换结果</returns>
        public static ConversionResult ConvertDocuments(
            string inputDirectory,
            string outputDirectory,
            WdSaveFormat targetFormat,
            string filePattern = "*.doc")
        {
            var result = new ConversionResult();

            try
            {
                // 确保输出目录存在
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                // 获取所有匹配的文件
                var files = Directory.GetFiles(inputDirectory, filePattern);

                Console.WriteLine($"找到 {files.Length} 个文档需要转换");

                result.TotalFiles = files.Length;
                result.ConvertedFiles = new List<string>();
                result.FailedFiles = new List<string>();

                foreach (var file in files)
                {
                    try
                    {
                        Console.WriteLine($"正在转换: {Path.GetFileName(file)}");

                        // 转换单个文档
                        ConvertSingleDocument(file, outputDirectory, targetFormat);

                        result.ConvertedFiles.Add(file);
                        Console.WriteLine($"转换完成: {Path.GetFileName(file)}");
                    }
                    catch (Exception ex)
                    {
                        result.FailedFiles.Add(file);
                        Console.WriteLine($"转换 {Path.GetFileName(file)} 时出错: {ex.Message}");
                    }
                }

                Console.WriteLine($"\n批量转换完成:");
                Console.WriteLine($"成功转换: {result.ConvertedFiles.Count} 个文档");
                Console.WriteLine($"转换失败: {result.FailedFiles.Count} 个文档");
                Console.WriteLine($"总计转换: {files.Length} 个文档");

                result.Success = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"批量转换过程中出错: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 转换单个文档
        /// </summary>
        /// <param name="inputFilePath">输入文件路径</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="targetFormat">目标格式</param>
        private static void ConvertSingleDocument(
            string inputFilePath,
            string outputDirectory,
            WdSaveFormat targetFormat)
        {
            using var app = WordFactory.Open(inputFilePath);
            var document = app.ActiveDocument;

            // 生成输出文件路径
            var fileName = Path.GetFileNameWithoutExtension(inputFilePath);
            string extension = GetExtensionForFormat(targetFormat);
            var outputFilePath = Path.Combine(outputDirectory, $"{fileName}{extension}");

            // 保存为指定格式
            document.SaveAs2(outputFilePath, targetFormat);

            Console.WriteLine($"  - 已转换为: {extension}");
        }

        /// <summary>
        /// 根据格式获取文件扩展名
        /// </summary>
        /// <param name="format">保存格式</param>
        /// <returns>文件扩展名</returns>
        private static string GetExtensionForFormat(WdSaveFormat format)
        {
            return format switch
            {
                WdSaveFormat.wdFormatDocument => ".doc",
                WdSaveFormat.wdFormatXMLDocument => ".xml",
                WdSaveFormat.wdFormatPDF => ".pdf",
                WdSaveFormat.wdFormatRTF => ".rtf",
                WdSaveFormat.wdFormatFilteredHTML => ".htm",
                WdSaveFormat.wdFormatHTML => ".html",
                _ => ".docx"
            };
        }

        /// <summary>
        /// 特殊转换示例：Word到PDF
        /// </summary>
        /// <param name="inputDirectory">输入目录</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>转换结果</returns>
        public static ConversionResult ConvertToPdf(string inputDirectory, string outputDirectory)
        {
            return ConvertDocuments(inputDirectory, outputDirectory, WdSaveFormat.wdFormatPDF, "*.docx");
        }

        /// <summary>
        /// 特殊转换示例：Word到HTML
        /// </summary>
        /// <param name="inputDirectory">输入目录</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>转换结果</returns>
        public static ConversionResult ConvertToHtml(string inputDirectory, string outputDirectory)
        {
            return ConvertDocuments(inputDirectory, outputDirectory, WdSaveFormat.wdFormatFilteredHTML, "*.docx");
        }

        /// <summary>
        /// 批量转换多种格式
        /// </summary>
        /// <param name="inputDirectory">输入目录</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="formats">目标格式列表</param>
        /// <param name="filePattern">文件匹配模式</param>
        /// <returns>转换结果</returns>
        public static MultiFormatConversionResult ConvertToMultipleFormats(
            string inputDirectory,
            string outputDirectory,
            List<WdSaveFormat> formats,
            string filePattern = "*.docx")
        {
            var result = new MultiFormatConversionResult();

            try
            {
                // 确保输出目录存在
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                // 获取所有匹配的文件
                var files = Directory.GetFiles(inputDirectory, filePattern);

                Console.WriteLine($"找到 {files.Length} 个文档需要转换");

                result.TotalFiles = files.Length;

                foreach (var format in formats)
                {
                    var formatResult = new ConversionResult
                    {
                        TotalFiles = files.Length
                    };

                    string formatName = GetFormatName(format);
                    Console.WriteLine($"\n开始转换为 {formatName} 格式:");

                    string formatOutputDirectory = Path.Combine(outputDirectory, formatName);
                    if (!Directory.Exists(formatOutputDirectory))
                    {
                        Directory.CreateDirectory(formatOutputDirectory);
                    }

                    foreach (var file in files)
                    {
                        try
                        {
                            Console.WriteLine($"  正在转换: {Path.GetFileName(file)}");

                            // 转换单个文档
                            ConvertSingleDocument(file, formatOutputDirectory, format);

                            formatResult.ConvertedFiles.Add(file);
                            Console.WriteLine($"  转换完成: {Path.GetFileName(file)}");
                        }
                        catch (Exception ex)
                        {
                            formatResult.FailedFiles.Add(file);
                            Console.WriteLine($"  转换 {Path.GetFileName(file)} 时出错: {ex.Message}");
                        }
                    }

                    result.FormatResults.Add(formatName, formatResult);

                    Console.WriteLine($"  {formatName} 格式转换完成:");
                    Console.WriteLine($"    成功转换: {formatResult.ConvertedFiles.Count} 个文档");
                    Console.WriteLine($"    转换失败: {formatResult.FailedFiles.Count} 个文档");
                }

                result.Success = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"批量多格式转换过程中出错: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 获取格式名称
        /// </summary>
        /// <param name="format">保存格式</param>
        /// <returns>格式名称</returns>
        private static string GetFormatName(WdSaveFormat format)
        {
            return format switch
            {
                WdSaveFormat.wdFormatDocument => "DOC",
                WdSaveFormat.wdFormatDocumentDefault => "DOCX",
                WdSaveFormat.wdFormatPDF => "PDF",
                WdSaveFormat.wdFormatRTF => "RTF",
                WdSaveFormat.wdFormatFilteredHTML => "HTML",
                WdSaveFormat.wdFormatXMLDocument => "XML",
                WdSaveFormat.wdFormatFlatXML => "FlatXML",
                _ => format.ToString()
            };
        }

        /// <summary>
        /// 异步转换文档
        /// </summary>
        /// <param name="inputDirectory">输入目录</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="targetFormat">目标格式</param>
        /// <param name="filePattern">文件匹配模式</param>
        /// <returns>转换结果</returns>
        public static async Task<ConversionResult> ConvertDocumentsAsync(
            string inputDirectory,
            string outputDirectory,
            WdSaveFormat targetFormat,
            string filePattern = "*.doc")
        {
            return await Task.Run(() => ConvertDocuments(inputDirectory, outputDirectory, targetFormat, filePattern));
        }

        /// <summary>
        /// 生成转换报告
        /// </summary>
        /// <param name="result">转换结果</param>
        /// <returns>转换报告</returns>
        public static string GenerateConversionReport(ConversionResult result)
        {
            var report = new StringBuilder();
            report.AppendLine("=== 文档转换报告 ===");
            report.AppendLine($"转换时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"总文件数: {result.TotalFiles}");
            report.AppendLine($"成功转换: {result.ConvertedFiles.Count}");
            report.AppendLine($"转换失败: {result.FailedFiles.Count}");
            report.AppendLine($"成功率: {(result.TotalFiles > 0 ? (double)result.ConvertedFiles.Count / result.TotalFiles * 100 : 0):F2}%");

            if (result.FailedFiles.Any())
            {
                report.AppendLine("\n失败文件列表:");
                foreach (var file in result.FailedFiles)
                {
                    report.AppendLine($"  - {Path.GetFileName(file)}");
                }
            }

            if (!string.IsNullOrEmpty(result.ErrorMessage))
            {
                report.AppendLine($"\n错误信息: {result.ErrorMessage}");
            }

            report.AppendLine("==================");

            return report.ToString();
        }

        /// <summary>
        /// 生成多格式转换报告
        /// </summary>
        /// <param name="result">多格式转换结果</param>
        /// <returns>转换报告</returns>
        public static string GenerateMultiFormatConversionReport(MultiFormatConversionResult result)
        {
            var report = new StringBuilder();
            report.AppendLine("=== 多格式文档转换报告 ===");
            report.AppendLine($"转换时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"总文件数: {result.TotalFiles}");

            foreach (var kvp in result.FormatResults)
            {
                var formatName = kvp.Key;
                var formatResult = kvp.Value;

                report.AppendLine($"\n{formatName} 格式转换结果:");
                report.AppendLine($"  成功转换: {formatResult.ConvertedFiles.Count}");
                report.AppendLine($"  转换失败: {formatResult.FailedFiles.Count}");
                report.AppendLine($"  成功率: {(formatResult.TotalFiles > 0 ? (double)formatResult.ConvertedFiles.Count / formatResult.TotalFiles * 100 : 0):F2}%");
            }

            if (!string.IsNullOrEmpty(result.ErrorMessage))
            {
                report.AppendLine($"\n错误信息: {result.ErrorMessage}");
            }

            report.AppendLine("========================");

            return report.ToString();
        }
    }

    /// <summary>
    /// 转换结果类
    /// </summary>
    public class ConversionResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 总文件数
        /// </summary>
        public int TotalFiles { get; set; }

        /// <summary>
        /// 已转换文件列表
        /// </summary>
        public List<string> ConvertedFiles { get; set; } = new List<string>();

        /// <summary>
        /// 失败文件列表
        /// </summary>
        public List<string> FailedFiles { get; set; } = new List<string>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 多格式转换结果类
    /// </summary>
    public class MultiFormatConversionResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 总文件数
        /// </summary>
        public int TotalFiles { get; set; }

        /// <summary>
        /// 各格式转换结果字典
        /// </summary>
        public Dictionary<string, ConversionResult> FormatResults { get; set; } = new Dictionary<string, ConversionResult>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }
}
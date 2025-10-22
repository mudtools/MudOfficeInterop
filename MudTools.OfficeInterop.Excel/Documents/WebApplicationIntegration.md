# 第18篇：Web应用集成方案详解

## 概述

Web应用集成是现代企业系统架构的重要组成部分。MudTools.OfficeInterop.Excel项目提供了完整的Web应用集成方案，支持服务器端Excel处理、文件上传下载、在线预览和权限控制等功能。本篇文章将详细介绍如何将Excel自动化功能集成到Web应用中。

## 服务器端Excel处理

### Web Excel处理服务

```csharp
using System;
using System.IO;
using System.Threading.Tasks;
using MudTools.OfficeInterop.Excel.CoreComponents.Core;

namespace MudTools.OfficeInterop.Excel.WebIntegration.Services
{
    /// <summary>
    /// Web Excel处理服务
    /// 提供服务器端的Excel文件处理功能
    /// </summary>
    public class WebExcelProcessingService : IDisposable
    {
        private IExcelApplication _application;
        private bool _disposed = false;
        
        public WebExcelProcessingService()
        {
            // 在服务器端创建Excel应用程序实例
            _application = ExcelFactory.Create();
            
            // 配置服务器端优化设置
            ConfigureServerSettings();
        }
        
        /// <summary>
        /// 配置服务器端设置
        /// </summary>
        private void ConfigureServerSettings()
        {
            // 禁用界面元素以提高性能
            _application.DisplayAlerts = false;
            _application.ScreenUpdating = false;
            _application.EnableEvents = false;
            
            // 设置计算模式
            _application.Calculation = CalculationMode.Manual;
        }
        
        /// <summary>
        /// 处理上传的Excel文件
        /// </summary>
        public async Task<ExcelProcessingResult> ProcessUploadedFile(Stream fileStream, string fileName)
        {
            var result = new ExcelProcessingResult(fileName);
            
            try
            {
                result.StartTime = DateTime.Now;
                
                // 将文件流保存到临时文件
                var tempFilePath = await SaveStreamToTempFile(fileStream, fileName);
                
                // 打开Excel文件
                var workbook = _application.OpenWorkbook(tempFilePath);
                
                // 执行处理逻辑
                await ProcessWorkbook(workbook, result);
                
                // 保存处理结果
                var outputPath = await SaveProcessedFile(workbook, fileName);
                result.OutputFilePath = outputPath;
                
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.Exception = ex;
            }
            finally
            {
                result.EndTime = DateTime.Now;
                result.Duration = result.EndTime - result.StartTime;
                
                // 清理临时文件
                CleanupTempFiles();
            }
            
            return result;
        }
        
        /// <summary>
        /// 将文件流保存到临时文件
        /// </summary>
        private async Task<string> SaveStreamToTempFile(Stream fileStream, string fileName)
        {
            var tempDir = Path.GetTempPath();
            var tempFileName = $"excel_upload_{Guid.NewGuid()}_{fileName}";
            var tempFilePath = Path.Combine(tempDir, tempFileName);
            
            using (var fileStreamOutput = File.Create(tempFilePath))
            {
                await fileStream.CopyToAsync(fileStreamOutput);
            }
            
            return tempFilePath;
        }
        
        /// <summary>
        /// 处理工作簿
        /// </summary>
        private async Task ProcessWorkbook(IExcelWorkbook workbook, ExcelProcessingResult result)
        {
            // 示例处理逻辑：数据验证和格式化
            foreach (var worksheet in workbook.Worksheets)
            {
                await ProcessWorksheet(worksheet, result);
            }
            
            // 重新计算公式
            _application.Calculate();
        }
        
        /// <summary>
        /// 处理工作表
        /// </summary>
        private async Task ProcessWorksheet(IExcelWorksheet worksheet, ExcelProcessingResult result)
        {
            // 处理数据验证
            await ValidateWorksheetData(worksheet, result);
            
            // 应用格式设置
            await ApplyWorksheetFormatting(worksheet, result);
            
            // 记录处理统计
            result.ProcessedWorksheets++;
        }
        
        /// <summary>
        /// 验证工作表数据
        /// </summary>
        private async Task ValidateWorksheetData(IExcelWorksheet worksheet, ExcelProcessingResult result)
        {
            var usedRange = worksheet.UsedRange;
            if (usedRange != null)
            {
                for (int row = 1; row <= usedRange.Rows.Count; row++)
                {
                    for (int col = 1; col <= usedRange.Columns.Count; col++)
                    {
                        var cell = usedRange.Cells[row, col];
                        if (cell != null && cell.Value != null)
                        {
                            // 简单的数据验证逻辑
                            await ValidateCellData(cell, result);
                        }
                    }
                }
            }
        }
        
        /// <summary>
        /// 验证单元格数据
        /// </summary>
        private async Task ValidateCellData(IExcelRange cell, ExcelProcessingResult result)
        {
            // 示例验证逻辑
            var value = cell.Value.ToString();
            
            if (string.IsNullOrWhiteSpace(value))
            {
                result.ValidationWarnings.Add($"空单元格: {cell.Address}");
            }
            else if (value.Length > 255)
            {
                result.ValidationWarnings.Add($"超长文本: {cell.Address}");
            }
        }
        
        /// <summary>
        /// 应用工作表格式
        /// </summary>
        private async Task ApplyWorksheetFormatting(IExcelWorksheet worksheet, ExcelProcessingResult result)
        {
            // 示例格式设置逻辑
            var usedRange = worksheet.UsedRange;
            if (usedRange != null)
            {
                // 设置表头格式
                if (usedRange.Rows.Count > 0)
                {
                    var headerRow = usedRange.Rows[1];
                    headerRow.Font.Bold = true;
                    headerRow.Interior.Color = System.Drawing.Color.LightGray.ToArgb();
                }
                
                // 设置数据区域格式
                if (usedRange.Rows.Count > 1)
                {
                    var dataRange = worksheet.Range[$"A2:{usedRange.Address.Split(':')[1]}"];
                    if (dataRange != null)
                    {
                        // 设置交替行颜色
                        ApplyAlternatingRowColors(dataRange);
                    }
                }
            }
        }
        
        /// <summary>
        /// 应用交替行颜色
        /// </summary>
        private void ApplyAlternatingRowColors(IExcelRange dataRange)
        {
            for (int row = 1; row <= dataRange.Rows.Count; row++)
            {
                var rowRange = dataRange.Rows[row];
                if (row % 2 == 0)
                {
                    // 偶数行设置浅色背景
                    rowRange.Interior.Color = System.Drawing.Color.WhiteSmoke.ToArgb();
                }
                else
                {
                    // 奇数行设置白色背景
                    rowRange.Interior.Color = System.Drawing.Color.White.ToArgb();
                }
            }
        }
        
        /// <summary>
        /// 保存处理后的文件
        /// </summary>
        private async Task<string> SaveProcessedFile(IExcelWorkbook workbook, string originalFileName)
        {
            var outputDir = Path.Combine(Path.GetTempPath(), "processed_excel");
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            
            var outputFileName = $"processed_{Guid.NewGuid()}_{originalFileName}";
            var outputFilePath = Path.Combine(outputDir, outputFileName);
            
            workbook.SaveAs(outputFilePath);
            workbook.Close();
            
            return outputFilePath;
        }
        
        /// <summary>
        /// 清理临时文件
        /// </summary>
        private void CleanupTempFiles()
        {
            try
            {
                var tempDir = Path.GetTempPath();
                var tempFiles = Directory.GetFiles(tempDir, "excel_upload_*");
                
                foreach (var tempFile in tempFiles)
                {
                    try
                    {
                        File.Delete(tempFile);
                    }
                    catch
                    {
                        // 忽略删除失败的文件
                    }
                }
            }
            catch
            {
                // 忽略清理错误
            }
        }
        
        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _application?.Quit();
                    _application = null;
                }
                
                _disposed = true;
            }
        }
        
        ~WebExcelProcessingService()
        {
            Dispose(false);
        }
    }
    
    /// <summary>
    /// Excel处理结果类
    /// </summary>
    public class ExcelProcessingResult
    {
        public string FileName { get; }
        public bool Success { get; set; }
        public string OutputFilePath { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan Duration { get; set; }
        public string ErrorMessage { get; set; }
        public Exception Exception { get; set; }
        public int ProcessedWorksheets { get; set; }
        public List<string> ValidationWarnings { get; set; }
        
        public ExcelProcessingResult(string fileName)
        {
            FileName = fileName;
            ValidationWarnings = new List<string>();
        }
    }
    
    /// <summary>
    /// 计算模式枚举
    /// </summary>
    public enum CalculationMode
    {
        Automatic,
        Manual
    }
}
```

### 批量处理服务

```csharp
/// <summary>
/// 批量Excel处理服务
/// 支持同时处理多个Excel文件
/// </summary>
public class BatchExcelProcessingService
{
    private readonly WebExcelProcessingService _processingService;
    private readonly int _maxConcurrentProcesses;
    
    public BatchExcelProcessingService(int maxConcurrentProcesses = 5)
    {
        _processingService = new WebExcelProcessingService();
        _maxConcurrentProcesses = maxConcurrentProcesses;
    }
    
    /// <summary>
    /// 批量处理Excel文件
    /// </summary>
    public async Task<BatchProcessingResult> ProcessBatchFiles(List<ExcelFileInfo> fileInfos)
    {
        var result = new BatchProcessingResult();
        
        try
        {
            result.StartTime = DateTime.Now;
            
            // 使用信号量控制并发数量
            var semaphore = new SemaphoreSlim(_maxConcurrentProcesses);
            var tasks = new List<Task<ExcelProcessingResult>>();
            
            foreach (var fileInfo in fileInfos)
            {
                await semaphore.WaitAsync();
                
                var task = Task.Run(async () =>
                {
                    try
                    {
                        return await _processingService.ProcessUploadedFile(fileInfo.FileStream, fileInfo.FileName);
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                });
                
                tasks.Add(task);
            }
            
            // 等待所有任务完成
            var results = await Task.WhenAll(tasks);
            
            // 收集结果
            foreach (var processingResult in results)
            {
                result.FileResults.Add(processingResult);
            }
            
            result.Success = result.FileResults.All(r => r.Success);
            result.TotalFilesProcessed = result.FileResults.Count;
            result.SuccessfulFiles = result.FileResults.Count(r => r.Success);
            result.FailedFiles = result.FileResults.Count(r => !r.Success);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
            result.Exception = ex;
        }
        finally
        {
            result.EndTime = DateTime.Now;
            result.Duration = result.EndTime - result.StartTime;
        }
        
        return result;
    }
    
    /// <summary>
    /// 处理特定类型的Excel文件
    /// </summary>
    public async Task<BatchProcessingResult> ProcessFilesByType(List<ExcelFileInfo> fileInfos, ExcelFileType fileType)
    {
        var filteredFiles = fileInfos.Where(f => f.FileType == fileType).ToList();
        return await ProcessBatchFiles(filteredFiles);
    }
    
    /// <summary>
    /// 获取处理统计信息
    /// </summary>
    public ProcessingStatistics GetProcessingStatistics(BatchProcessingResult result)
    {
        var stats = new ProcessingStatistics
        {
            TotalFiles = result.TotalFilesProcessed,
            SuccessfulFiles = result.SuccessfulFiles,
            FailedFiles = result.FailedFiles,
            TotalProcessingTime = result.Duration,
            AverageProcessingTime = TimeSpan.FromTicks(
                result.FileResults.Sum(r => r.Duration.Ticks) / Math.Max(1, result.FileResults.Count))
        };
        
        // 计算成功率
        stats.SuccessRate = result.TotalFilesProcessed > 0 
            ? (double)result.SuccessfulFiles / result.TotalFilesProcessed * 100 
            : 0;
        
        return stats;
    }
    
    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        _processingService?.Dispose();
    }
}

/// <summary>
/// Excel文件信息类
/// </summary>
public class ExcelFileInfo
{
    public Stream FileStream { get; set; }
    public string FileName { get; set; }
    public ExcelFileType FileType { get; set; }
    public long FileSize { get; set; }
    public DateTime UploadTime { get; set; }
    public Dictionary<string, object> Metadata { get; set; }
    
    public ExcelFileInfo()
    {
        Metadata = new Dictionary<string, object>();
        UploadTime = DateTime.Now;
    }
}

/// <summary>
/// Excel文件类型枚举
/// </summary>
public enum ExcelFileType
{
    Report,         // 报表文件
    Template,       // 模板文件
    DataImport,     // 数据导入文件
    Export,         // 导出文件
    Unknown         // 未知类型
}

/// <summary>
/// 批量处理结果类
/// </summary>
public class BatchProcessingResult
{
    public bool Success { get; set; }
    public string ErrorMessage { get; set; }
    public Exception Exception { get; set; }
    public DateTime StartTime { get; set; }
    public DateTime EndTime { get; set; }
    public TimeSpan Duration { get; set; }
    public int TotalFilesProcessed { get; set; }
    public int SuccessfulFiles { get; set; }
    public int FailedFiles { get; set; }
    public List<ExcelProcessingResult> FileResults { get; set; }
    
    public BatchProcessingResult()
    {
        FileResults = new List<ExcelProcessingResult>();
    }
}

/// <summary>
/// 处理统计信息类
/// </summary>
public class ProcessingStatistics
{
    public int TotalFiles { get; set; }
    public int SuccessfulFiles { get; set; }
    public int FailedFiles { get; set; }
    public double SuccessRate { get; set; }
    public TimeSpan TotalProcessingTime { get; set; }
    public TimeSpan AverageProcessingTime { get; set; }
}
```

## 文件上传下载功能

### Web文件管理器

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;

namespace MudTools.OfficeInterop.Excel.WebIntegration.FileManagement
{
    /// <summary>
    /// Web文件管理器
    /// 提供Excel文件的上传、下载和管理功能
    /// </summary>
    public class WebFileManager
    {
        private readonly string _uploadDirectory;
        private readonly string _downloadDirectory;
        private readonly long _maxFileSize;
        private readonly List<string> _allowedExtensions;
        
        public WebFileManager(string uploadDirectory, string downloadDirectory, 
            long maxFileSize = 10 * 1024 * 1024) // 默认10MB
        {
            _uploadDirectory = uploadDirectory ?? throw new ArgumentNullException(nameof(uploadDirectory));
            _downloadDirectory = downloadDirectory ?? throw new ArgumentNullException(nameof(downloadDirectory));
            _maxFileSize = maxFileSize;
            
            _allowedExtensions = new List<string>
            {
                ".xlsx", ".xls", ".xlsm", ".xltx", ".xltm", ".csv"
            };
            
            // 确保目录存在
            EnsureDirectoryExists(_uploadDirectory);
            EnsureDirectoryExists(_downloadDirectory);
        }
        
        /// <summary>
        /// 确保目录存在
        /// </summary>
        private void EnsureDirectoryExists(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
        }
        
        /// <summary>
        /// 上传Excel文件
        /// </summary>
        public async Task<FileUploadResult> UploadExcelFile(Stream fileStream, string fileName, 
            string userId, Dictionary<string, object> metadata = null)
        {
            var result = new FileUploadResult(fileName, userId);
            
            try
            {
                result.StartTime = DateTime.Now;
                
                // 验证文件
                var validationResult = ValidateFile(fileStream, fileName);
                if (!validationResult.IsValid)
                {
                    result.Success = false;
                    result.ErrorMessage = validationResult.ErrorMessage;
                    return result;
                }
                
                // 生成唯一文件名
                var uniqueFileName = GenerateUniqueFileName(fileName, userId);
                var filePath = Path.Combine(_uploadDirectory, uniqueFileName);
                
                // 保存文件
                using (var fileStreamOutput = File.Create(filePath))
                {
                    await fileStream.CopyToAsync(fileStreamOutput);
                }
                
                // 计算文件哈希
                result.FileHash = await CalculateFileHash(filePath);
                result.FileSize = new FileInfo(filePath).Length;
                result.FilePath = filePath;
                result.UniqueFileName = uniqueFileName;
                result.Metadata = metadata ?? new Dictionary<string, object>();
                
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.Exception = ex;
            }
            finally
            {
                result.EndTime = DateTime.Now;
                result.Duration = result.EndTime - result.StartTime;
            }
            
            return result;
        }
        
        /// <summary>
        /// 验证文件
        /// </summary>
        private FileValidationResult ValidateFile(Stream fileStream, string fileName)
        {
            var result = new FileValidationResult();
            
            // 检查文件扩展名
            var extension = Path.GetExtension(fileName)?.ToLower();
            if (string.IsNullOrEmpty(extension) || !_allowedExtensions.Contains(extension))
            {
                result.IsValid = false;
                result.ErrorMessage = $"不支持的文件格式: {extension}";
                return result;
            }
            
            // 检查文件大小
            if (fileStream.Length > _maxFileSize)
            {
                result.IsValid = false;
                result.ErrorMessage = $"文件大小超过限制: {fileStream.Length} > {_maxFileSize}";
                return result;
            }
            
            // 检查文件内容（简化实现）
            if (!IsValidExcelFile(fileStream, extension))
            {
                result.IsValid = false;
                result.ErrorMessage = "文件内容无效或已损坏";
                return result;
            }
            
            result.IsValid = true;
            return result;
        }
        
        /// <summary>
        /// 检查是否为有效的Excel文件
        /// </summary>
        private bool IsValidExcelFile(Stream fileStream, string extension)
        {
            // 简化实现，实际应用中应进行更严格的验证
            return fileStream.Length > 0;
        }
        
        /// <summary>
        /// 生成唯一文件名
        /// </summary>
        private string GenerateUniqueFileName(string originalFileName, string userId)
        {
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalFileName);
            var extension = Path.GetExtension(originalFileName);
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var randomPart = Guid.NewGuid().ToString("N").Substring(0, 8);
            
            return $"{fileNameWithoutExtension}_{userId}_{timestamp}_{randomPart}{extension}";
        }
        
        /// <summary>
        /// 计算文件哈希
        /// </summary>
        private async Task<string> CalculateFileHash(string filePath)
        {
            using (var md5 = MD5.Create())
            using (var stream = File.OpenRead(filePath))
            {
                var hash = await md5.ComputeHashAsync(stream);
                return BitConverter.ToString(hash).Replace("-", "").ToLower();
            }
        }
        
        /// <summary>
        /// 下载Excel文件
        /// </summary>
        public async Task<FileDownloadResult> DownloadExcelFile(string fileId, string userId)
        {
            var result = new FileDownloadResult(fileId, userId);
            
            try
            {
                result.StartTime = DateTime.Now;
                
                // 查找文件
                var fileInfo = await FindFileById(fileId, userId);
                if (fileInfo == null)
                {
                    result.Success = false;
                    result.ErrorMessage = "文件不存在或无权访问";
                    return result;
                }
                
                // 验证访问权限
                if (!HasDownloadPermission(fileInfo, userId))
                {
                    result.Success = false;
                    result.ErrorMessage = "无权下载此文件";
                    return result;
                }
                
                // 读取文件内容
                result.FileContent = await File.ReadAllBytesAsync(fileInfo.FilePath);
                result.FileName = fileInfo.OriginalFileName;
                result.ContentType = GetContentType(fileInfo.FilePath);
                
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.Exception = ex;
            }
            finally
            {
                result.EndTime = DateTime.Now;
                result.Duration = result.EndTime - result.StartTime;
            }
            
            return result;
        }
        
        /// <summary>
        /// 根据文件ID查找文件
        /// </summary>
        private async Task<StoredFileInfo> FindFileById(string fileId, string userId)
        {
            // 简化实现，实际应用中应从数据库查询
            var files = Directory.GetFiles(_uploadDirectory, $"*{userId}*");
            var filePath = files.FirstOrDefault(f => f.Contains(fileId));
            
            if (filePath != null && File.Exists(filePath))
            {
                return new StoredFileInfo
                {
                    FilePath = filePath,
                    OriginalFileName = Path.GetFileName(filePath),
                    UploadTime = File.GetCreationTime(filePath),
                    FileSize = new FileInfo(filePath).Length
                };
            }
            
            return null;
        }
        
        /// <summary>
        /// 检查下载权限
        /// </summary>
        private bool HasDownloadPermission(StoredFileInfo fileInfo, string userId)
        {
            // 简化实现，实际应用中应进行更复杂的权限检查
            return fileInfo.FilePath.Contains(userId);
        }
        
        /// <summary>
        /// 获取内容类型
        /// </summary>
        private string GetContentType(string filePath)
        {
            var extension = Path.GetExtension(filePath)?.ToLower();
            
            return extension switch
            {
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".xls" => "application/vnd.ms-excel",
                ".xlsm" => "application/vnd.ms-excel.sheet.macroEnabled.12",
                ".xltx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
                ".xltm" => "application/vnd.ms-excel.template.macroEnabled.12",
                ".csv" => "text/csv",
                _ => "application/octet-stream"
            };
        }
        
        /// <summary>
        /// 获取用户文件列表
        /// </summary>
        public async Task<List<StoredFileInfo>> GetUserFiles(string userId, int page = 1, int pageSize = 20)
        {
            var files = Directory.GetFiles(_uploadDirectory, $"*{userId}*")
                .OrderByDescending(f => File.GetCreationTime(f))
                .Skip((page - 1) * pageSize)
                .Take(pageSize);
            
            var fileInfos = new List<StoredFileInfo>();
            
            foreach (var filePath in files)
            {
                var fileInfo = new FileInfo(filePath);
                fileInfos.Add(new StoredFileInfo
                {
                    FilePath = filePath,
                    OriginalFileName = Path.GetFileName(filePath),
                    UploadTime = fileInfo.CreationTime,
                    FileSize = fileInfo.Length
                });
            }
            
            return await Task.FromResult(fileInfos);
        }
        
        /// <summary>
        /// 删除文件
        /// </summary>
        public async Task<FileDeleteResult> DeleteFile(string fileId, string userId)
        {
            var result = new FileDeleteResult(fileId, userId);
            
            try
            {
                result.StartTime = DateTime.Now;
                
                // 查找文件
                var fileInfo = await FindFileById(fileId, userId);
                if (fileInfo == null)
                {
                    result.Success = false;
                    result.ErrorMessage = "文件不存在";
                    return result;
                }
                
                // 验证删除权限
                if (!HasDeletePermission(fileInfo, userId))
                {
                    result.Success = false;
                    result.ErrorMessage = "无权删除此文件";
                    return result;
                }
                
                // 删除文件
                File.Delete(fileInfo.FilePath);
                
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.Exception = ex;
            }
            finally
            {
                result.EndTime = DateTime.Now;
                result.Duration = result.EndTime - result.StartTime;
            }
            
            return result;
        }
        
        /// <summary>
        /// 检查删除权限
        /// </summary>
        private bool HasDeletePermission(StoredFileInfo fileInfo, string userId)
        {
            // 简化实现
            return fileInfo.FilePath.Contains(userId);
        }
    }
    
    /// <summary>
    /// 文件上传结果类
    /// </summary>
    public class FileUploadResult
    {
        public string FileName { get; }
        public string UserId { get; }
        public bool Success { get; set; }
        public string FilePath { get; set; }
        public string UniqueFileName { get; set; }
        public string FileHash { get; set; }
        public long FileSize { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan Duration { get; set; }
        public string ErrorMessage { get; set; }
        public Exception Exception { get; set; }
        public Dictionary<string, object> Metadata { get; set; }
        
        public FileUploadResult(string fileName, string userId)
        {
            FileName = fileName;
            UserId = userId;
            Metadata = new Dictionary<string, object>();
        }
    }
    
    /// <summary>
    /// 文件验证结果类
    /// </summary>
    public class FileValidationResult
    {
        public bool IsValid { get; set; }
        public string ErrorMessage { get; set; }
    }
    
    /// <summary>
    /// 存储的文件信息类
    /// </summary>
    public class StoredFileInfo
    {
        public string FilePath { get; set; }
        public string OriginalFileName { get; set; }
        public DateTime UploadTime { get; set; }
        public long FileSize { get; set; }
    }
    
    /// <summary>
    /// 文件下载结果类
    /// </summary>
    public class FileDownloadResult
    {
        public string FileId { get; }
        public string UserId { get; }
        public bool Success { get; set; }
        public byte[] FileContent { get; set; }
        public string FileName { get; set; }
        public string ContentType { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan Duration { get; set; }
        public string ErrorMessage { get; set; }
        public Exception Exception { get; set; }
        
        public FileDownloadResult(string fileId, string userId)
        {
            FileId = fileId;
            UserId = userId;
        }
    }
    
    /// <summary>
    /// 文件删除结果类
    /// </summary>
    public class FileDeleteResult
    {
        public string FileId { get; }
        public string UserId { get; }
        public bool Success { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan Duration { get; set; }
        public string ErrorMessage { get; set; }
        public Exception Exception { get; set; }
        
        public FileDeleteResult(string fileId, string userId)
        {
            FileId = fileId;
            UserId = userId;
        }
    }
}
```

## 在线预览功能

### Excel预览服务

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using MudTools.OfficeInterop.Excel.CoreComponents.Core;

namespace MudTools.OfficeInterop.Excel.WebIntegration.Preview
{
    /// <summary>
    /// Excel预览服务
    /// 提供Excel文件的在线预览功能
    /// </summary>
    public class ExcelPreviewService : IDisposable
    {
        private readonly WebExcelProcessingService _processingService;
        private readonly WebFileManager _fileManager;
        private bool _disposed = false;
        
        public ExcelPreviewService(WebFileManager fileManager)
        {
            _processingService = new WebExcelProcessingService();
            _fileManager = fileManager ?? throw new ArgumentNullException(nameof(fileManager));
        }
        
        /// <summary>
        /// 生成Excel预览
        /// </summary>
        public async Task<ExcelPreviewResult> GeneratePreview(string fileId, string userId, PreviewOptions options = null)
        {
            var result = new ExcelPreviewResult(fileId, userId);
            
            try
            {
                result.StartTime = DateTime.Now;
                
                // 下载文件
                var downloadResult = await _fileManager.DownloadExcelFile(fileId, userId);
                if (!downloadResult.Success)
                {
                    result.Success = false;
                    result.ErrorMessage = $"文件下载失败: {downloadResult.ErrorMessage}";
                    return result;
                }
                
                // 将文件内容保存到临时文件
                var tempFilePath = await SaveContentToTempFile(downloadResult.FileContent, downloadResult.FileName);
                
                // 生成预览
                await GeneratePreviewFromFile(tempFilePath, result, options);
                
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.Exception = ex;
            }
            finally
            {
                result.EndTime = DateTime.Now;
                result.Duration = result.EndTime - result.StartTime;
                
                // 清理临时文件
                CleanupTempFiles();
            }
            
            return result;
        }
        
        /// <summary>
        /// 将文件内容保存到临时文件
        /// </summary>
        private async Task<string> SaveContentToTempFile(byte[] fileContent, string fileName)
        {
            var tempDir = Path.GetTempPath();
            var tempFileName = $"preview_{Guid.NewGuid()}_{fileName}";
            var tempFilePath = Path.Combine(tempDir, tempFileName);
            
            await File.WriteAllBytesAsync(tempFilePath, fileContent);
            
            return tempFilePath;
        }
        
        /// <summary>
        /// 从文件生成预览
        /// </summary>
        private async Task GeneratePreviewFromFile(string filePath, ExcelPreviewResult result, PreviewOptions options)
        {
            options ??= new PreviewOptions();
            
            // 打开Excel文件
            var workbook = _processingService.OpenWorkbook(filePath);
            
            // 生成预览数据
            await GeneratePreviewData(workbook, result, options);
            
            // 生成预览图像（如果启用）
            if (options.GenerateImages)
            {
                await GeneratePreviewImages(workbook, result, options);
            }
            
            workbook.Close();
        }
        
        /// <summary>
        /// 生成预览数据
        /// </summary>
        private async Task GeneratePreviewData(IExcelWorkbook workbook, ExcelPreviewResult result, PreviewOptions options)
        {
            result.WorksheetPreviews = new List<WorksheetPreview>();
            
            foreach (var worksheet in workbook.Worksheets)
            {
                var worksheetPreview = await GenerateWorksheetPreview(worksheet, options);
                result.WorksheetPreviews.Add(worksheetPreview);
            }
            
            // 生成工作簿摘要
            result.WorkbookSummary = GenerateWorkbookSummary(workbook);
        }
        
        /// <summary>
        /// 生成工作表预览
        /// </summary>
        private async Task<WorksheetPreview> GenerateWorksheetPreview(IExcelWorksheet worksheet, PreviewOptions options)
        {
            var preview = new WorksheetPreview
            {
                Name = worksheet.Name,
                PreviewData = new List<List<object>>()
            };
            
            var usedRange = worksheet.UsedRange;
            if (usedRange != null)
            {
                // 限制预览行数和列数
                var maxRows = Math.Min(options.MaxPreviewRows, usedRange.Rows.Count);
                var maxCols = Math.Min(options.MaxPreviewColumns, usedRange.Columns.Count);
                
                // 提取预览数据
                for (int row = 1; row <= maxRows; row++)
                {
                    var rowData = new List<object>();
                    
                    for (int col = 1; col <= maxCols; col++)
                    {
                        var cell = usedRange.Cells[row, col];
                        rowData.Add(cell?.Value ?? string.Empty);
                    }
                    
                    preview.PreviewData.Add(rowData);
                }
                
                // 生成工作表统计信息
                preview.Statistics = GenerateWorksheetStatistics(worksheet, usedRange);
            }
            
            return preview;
        }
        
        /// <summary>
        /// 生成工作表统计信息
        /// </summary>
        private WorksheetStatistics GenerateWorksheetStatistics(IExcelWorksheet worksheet, IExcelRange usedRange)
        {
            return new WorksheetStatistics
            {
                RowCount = usedRange.Rows.Count,
                ColumnCount = usedRange.Columns.Count,
                CellCount = usedRange.Rows.Count * usedRange.Columns.Count,
                HasFormulas = CheckHasFormulas(worksheet),
                HasCharts = worksheet.Charts.Count > 0,
                HasPivotTables = worksheet.PivotTables.Count > 0
            };
        }
        
        /// <summary>
        /// 检查是否有公式
        /// </summary>
        private bool CheckHasFormulas(IExcelWorksheet worksheet)
        {
            var usedRange = worksheet.UsedRange;
            if (usedRange != null)
            {
                for (int row = 1; row <= usedRange.Rows.Count; row++)
                {
                    for (int col = 1; col <= usedRange.Columns.Count; col++)
                    {
                        var cell = usedRange.Cells[row, col];
                        if (cell != null && !string.IsNullOrEmpty(cell.Formula))
                        {
                            return true;
                        }
                    }
                }
            }
            
            return false;
        }
        
        /// <summary>
        /// 生成工作簿摘要
        /// </summary>
        private WorkbookSummary GenerateWorkbookSummary(IExcelWorkbook workbook)
        {
            return new WorkbookSummary
            {
                WorksheetCount = workbook.Worksheets.Count,
                TotalCells = CalculateTotalCells(workbook),
                FileSize = 0, // 需要从文件系统获取
                CreatedDate = DateTime.Now, // 需要从文件属性获取
                ModifiedDate = DateTime.Now // 需要从文件属性获取
            };
        }
        
        /// <summary>
        /// 计算总单元格数
        /// </summary>
        private int CalculateTotalCells(IExcelWorkbook workbook)
        {
            int totalCells = 0;
            
            foreach (var worksheet in workbook.Worksheets)
            {
                var usedRange = worksheet.UsedRange;
                if (usedRange != null)
                {
                    totalCells += usedRange.Rows.Count * usedRange.Columns.Count;
                }
            }
            
            return totalCells;
        }
        
        /// <summary>
        /// 生成预览图像
        /// </summary>
        private async Task GeneratePreviewImages(IExcelWorkbook workbook, ExcelPreviewResult result, PreviewOptions options)
        {
            result.PreviewImages = new List<PreviewImage>();
            
            foreach (var worksheet in workbook.Worksheets)
            {
                if (result.WorksheetPreviews.Count < options.MaxWorksheetsForImages)
                {
                    var image = await GenerateWorksheetImage(worksheet, options);
                    if (image != null)
                    {
                        result.PreviewImages.Add(image);
                    }
                }
            }
        }
        
        /// <summary>
        /// 生成工作表图像
        /// </summary>
        private async Task<PreviewImage> GenerateWorksheetImage(IExcelWorksheet worksheet, PreviewOptions options)
        {
            // 简化实现，实际应用中应使用Excel的导出图像功能
            return await Task.FromResult<PreviewImage>(null);
        }
        
        /// <summary>
        /// 清理临时文件
        /// </summary>
        private void CleanupTempFiles()
        {
            try
            {
                var tempDir = Path.GetTempPath();
                var tempFiles = Directory.GetFiles(tempDir, "preview_*");
                
                foreach (var tempFile in tempFiles)
                {
                    try
                    {
                        File.Delete(tempFile);
                    }
                    catch
                    {
                        // 忽略删除失败的文件
                    }
                }
            }
            catch
            {
                // 忽略清理错误
            }
        }
        
        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _processingService?.Dispose();
                }
                
                _disposed = true;
            }
        }
        
        ~ExcelPreviewService()
        {
            Dispose(false);
        }
    }
    
    /// <summary>
    /// Excel预览结果类
    /// </summary>
    public class ExcelPreviewResult
    {
        public string FileId { get; }
        public string UserId { get; }
        public bool Success { get; set; }
        public List<WorksheetPreview> WorksheetPreviews { get; set; }
        public WorkbookSummary WorkbookSummary { get; set; }
        public List<PreviewImage> PreviewImages { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan Duration { get; set; }
        public string ErrorMessage { get; set; }
        public Exception Exception { get; set; }
        
        public ExcelPreviewResult(string fileId, string userId)
        {
            FileId = fileId;
            UserId = userId;
            WorksheetPreviews = new List<WorksheetPreview>();
            PreviewImages = new List<PreviewImage>();
        }
    }
    
    /// <summary>
    /// 预览选项类
    /// </summary>
    public class PreviewOptions
    {
        public int MaxPreviewRows { get; set; } = 50;
        public int MaxPreviewColumns { get; set; } = 20;
        public bool GenerateImages { get; set; } = false;
        public int MaxWorksheetsForImages { get; set; } = 5;
        public PreviewImageFormat ImageFormat { get; set; } = PreviewImageFormat.PNG;
    }
    
    /// <summary>
    /// 工作表预览类
    /// </summary>
    public class WorksheetPreview
    {
        public string Name { get; set; }
        public List<List<object>> PreviewData { get; set; }
        public WorksheetStatistics Statistics { get; set; }
    }
    
    /// <summary>
    /// 工作表统计信息类
    /// </summary>
    public class WorksheetStatistics
    {
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public int CellCount { get; set; }
        public bool HasFormulas { get; set; }
        public bool HasCharts { get; set; }
        public bool HasPivotTables { get; set; }
    }
    
    /// <summary>
    /// 工作簿摘要类
    /// </summary>
    public class WorkbookSummary
    {
        public int WorksheetCount { get; set; }
        public int TotalCells { get; set; }
        public long FileSize { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime ModifiedDate { get; set; }
    }
    
    /// <summary>
    /// 预览图像类
    /// </summary>
    public class PreviewImage
    {
        public string WorksheetName { get; set; }
        public byte[] ImageData { get; set; }
        public string Format { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }
    
    /// <summary>
    /// 预览图像格式枚举
    /// </summary>
    public enum PreviewImageFormat
    {
        PNG,
        JPEG,
        GIF
    }
}
```

## 权限和安全考虑

### 安全权限管理器

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MudTools.OfficeInterop.Excel.WebIntegration.Security
{
    /// <summary>
    /// 安全权限管理器
    /// 提供Web应用的安全控制和权限管理
    /// </summary>
    public class SecurityPermissionManager
    {
        private readonly Dictionary<string, UserPermission> _userPermissions;
        private readonly List<SecurityRule> _securityRules;
        
        public SecurityPermissionManager()
        {
            _userPermissions = new Dictionary<string, UserPermission>();
            _securityRules = new List<SecurityRule>();
            
            InitializeDefaultRules();
        }
        
        /// <summary>
        /// 初始化默认安全规则
        /// </summary>
        private void InitializeDefaultRules()
        {
            // 文件大小限制规则
            _securityRules.Add(new SecurityRule
            {
                RuleType = SecurityRuleType.FileSizeLimit,
                Condition = (context) => context.FileSize <= 10 * 1024 * 1024, // 10MB
                Action = SecurityAction.Deny,
                Message = "文件大小超过限制"
            });
            
            // 文件类型限制规则
            _securityRules.Add(new SecurityRule
            {
                RuleType = SecurityRuleType.FileTypeRestriction,
                Condition = (context) => IsAllowedFileType(context.FileExtension),
                Action = SecurityAction.Deny,
                Message = "不支持的文件类型"
            });
            
            // 用户权限检查规则
            _securityRules.Add(new SecurityRule
            {
                RuleType = SecurityRuleType.UserPermission,
                Condition = (context) => HasUserPermission(context.UserId, context.Operation),
                Action = SecurityAction.Deny,
                Message = "用户无权执行此操作"
            });
        }
        
        /// <summary>
        /// 检查是否为允许的文件类型
        /// </summary>
        private bool IsAllowedFileType(string fileExtension)
        {
            var allowedExtensions = new[] { ".xlsx", ".xls", ".xlsm", ".xltx", ".xltm", ".csv" };
            return allowedExtensions.Contains(fileExtension?.ToLower());
        }
        
        /// <summary>
        /// 检查用户权限
        /// </summary>
        private bool HasUserPermission(string userId, SecurityOperation operation)
        {
            if (_userPermissions.TryGetValue(userId, out var permission))
            {
                return permission.AllowedOperations.Contains(operation);
            }
            
            // 默认权限
            return operation == SecurityOperation.Upload || operation == SecurityOperation.Download;
        }
        
        /// <summary>
        /// 验证安全权限
        /// </summary>
        public async Task<SecurityValidationResult> ValidateSecurity(SecurityContext context)
        {
            var result = new SecurityValidationResult(context);
            
            try
            {
                // 应用所有安全规则
                foreach (var rule in _securityRules)
                {
                    var ruleResult = await ApplySecurityRule(rule, context);
                    result.RuleResults.Add(ruleResult);
                    
                    if (ruleResult.Action == SecurityAction.Deny)
                    {
                        result.IsValid = false;
                        result.DenyReason = ruleResult.Message;
                        break;
                    }
                }
                
                if (result.IsValid)
                {
                    result.Message = "安全验证通过";
                }
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.DenyReason = $"安全验证异常: {ex.Message}";
                result.Exception = ex;
            }
            
            return result;
        }
        
        /// <summary>
        /// 应用安全规则
        /// </summary>
        private async Task<RuleValidationResult> ApplySecurityRule(SecurityRule rule, SecurityContext context)
        {
            var result = new RuleValidationResult(rule.RuleType);
            
            try
            {
                var conditionResult = await Task.Run(() => rule.Condition(context));
                
                if (conditionResult)
                {
                    result.Action = SecurityAction.Allow;
                    result.Message = $"规则'{rule.RuleType}'验证通过";
                }
                else
                {
                    result.Action = rule.Action;
                    result.Message = rule.Message;
                }
            }
            catch (Exception ex)
            {
                result.Action = SecurityAction.Deny;
                result.Message = $"规则'{rule.RuleType}'执行异常: {ex.Message}";
                result.Exception = ex;
            }
            
            return result;
        }
        
        /// <summary>
        /// 设置用户权限
        /// </summary>
        public void SetUserPermission(string userId, UserPermission permission)
        {
            _userPermissions[userId] = permission;
        }
        
        /// <summary>
        /// 添加安全规则
        /// </summary>
        public void AddSecurityRule(SecurityRule rule)
        {
            _securityRules.Add(rule);
        }
        
        /// <summary>
        /// 获取安全审计日志
        /// </summary>
        public async Task<List<SecurityAuditLog>> GetSecurityAuditLogs(string userId = null, 
            DateTime? startDate = null, DateTime? endDate = null)
        {
            // 简化实现，实际应用中应从数据库查询
            return await Task.FromResult(new List<SecurityAuditLog>());
        }
    }
    
    /// <summary>
    /// 用户权限类
    /// </summary>
    public class UserPermission
    {
        public string UserId { get; set; }
        public List<SecurityOperation> AllowedOperations { get; set; }
        public DateTime ValidFrom { get; set; }
        public DateTime ValidTo { get; set; }
        
        public UserPermission()
        {
            AllowedOperations = new List<SecurityOperation>();
            ValidFrom = DateTime.Now;
            ValidTo = DateTime.Now.AddYears(1);
        }
    }
    
    /// <summary>
    /// 安全规则类
    /// </summary>
    public class SecurityRule
    {
        public SecurityRuleType RuleType { get; set; }
        public Func<SecurityContext, bool> Condition { get; set; }
        public SecurityAction Action { get; set; }
        public string Message { get; set; }
    }
    
    /// <summary>
    /// 安全上下文类
    /// </summary>
    public class SecurityContext
    {
        public string UserId { get; set; }
        public SecurityOperation Operation { get; set; }
        public string FileName { get; set; }
        public string FileExtension { get; set; }
        public long FileSize { get; set; }
        public string IpAddress { get; set; }
        public DateTime RequestTime { get; set; }
        
        public SecurityContext()
        {
            RequestTime = DateTime.Now;
        }
    }
    
    /// <summary>
    /// 安全验证结果类
    /// </summary>
    public class SecurityValidationResult
    {
        public SecurityContext Context { get; }
        public bool IsValid { get; set; } = true;
        public string Message { get; set; }
        public string DenyReason { get; set; }
        public Exception Exception { get; set; }
        public List<RuleValidationResult> RuleResults { get; set; }
        
        public SecurityValidationResult(SecurityContext context)
        {
            Context = context;
            RuleResults = new List<RuleValidationResult>();
        }
    }
    
    /// <summary>
    /// 规则验证结果类
    /// </summary>
    public class RuleValidationResult
    {
        public SecurityRuleType RuleType { get; }
        public SecurityAction Action { get; set; }
        public string Message { get; set; }
        public Exception Exception { get; set; }
        
        public RuleValidationResult(SecurityRuleType ruleType)
        {
            RuleType = ruleType;
        }
    }
    
    /// <summary>
    /// 安全审计日志类
    /// </summary>
    public class SecurityAuditLog
    {
        public string UserId { get; set; }
        public SecurityOperation Operation { get; set; }
        public string FileName { get; set; }
        public SecurityAction Action { get; set; }
        public DateTime Timestamp { get; set; }
        public string IpAddress { get; set; }
        public string Details { get; set; }
    }
    
    /// <summary>
    /// 安全规则类型枚举
    /// </summary>
    public enum SecurityRuleType
    {
        FileSizeLimit,
        FileTypeRestriction,
        UserPermission,
        IpRestriction,
        TimeRestriction
    }
    
    /// <summary>
    /// 安全操作枚举
    /// </summary>
    public enum SecurityOperation
    {
        Upload,
        Download,
        Delete,
        Preview,
        Process,
        Admin
    }
    
    /// <summary>
    /// 安全动作枚举
    /// </summary>
    public enum SecurityAction
    {
        Allow,
        Deny,
        RequireApproval
    }
}
```

## 总结

本篇博文详细介绍了基于MudTools.OfficeInterop.Excel项目构建Web应用集成方案的完整方案，包括：

1. **服务器端Excel处理**：Web Excel处理服务、批量处理服务
2. **文件上传下载功能**：Web文件管理器、文件验证、权限控制
3. **在线预览功能**：Excel预览服务、预览数据生成、图像预览
4. **权限和安全考虑**：安全权限管理器、安全规则、审计日志

### 系统特色

**高性能服务器端处理**
- 优化的Excel应用程序配置
- 并发处理控制
- 内存和资源管理

**完整的文件管理**
- 安全的文件上传验证
- 灵活的文件下载机制
- 完善的文件生命周期管理

**丰富的预览功能**
- 多维度预览数据生成
- 图像预览支持
- 预览选项配置

**企业级安全控制**
- 多层次权限验证
- 可配置的安全规则
- 完整的审计日志

### 实际应用价值

通过本方案，企业可以实现：
- **Web端Excel自动化**：将Excel处理功能集成到Web应用中
- **安全的文件管理**：确保文件上传下载的安全性
- **在线协作**：支持多人协作处理Excel文档
- **移动办公**：移动端Excel文件处理和预览

这套Web应用集成方案为企业的现代化系统架构提供了强大的技术支撑，可以直接应用于实际的Web应用开发中。
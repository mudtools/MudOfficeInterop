using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace BestPracticesSummarySample
{
    /// <summary>
    /// Excel自动化开发最佳实践示例程序
    /// 演示Excel自动化开发中的核心技术和最佳实践
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("Excel自动化开发最佳实践示例");
            Console.WriteLine("========================");
            Console.WriteLine();

            // 演示项目架构最佳实践
            LayeredArchitectureExample();
            
            // 演示性能优化最佳实践
            PerformanceOptimizationExample();
            
            // 演示错误处理最佳实践
            ErrorHandlingExample();
            
            // 演示资源管理最佳实践
            ResourceManagementExample();
            
            // 演示安全开发最佳实践
            SecurityBestPracticesExample();
            
            // 演示代码组织最佳实践
            CodeOrganizationExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 项目架构最佳实践示例
        /// 演示如何设计良好的项目架构
        /// </summary>
        static void LayeredArchitectureExample()
        {
            Console.WriteLine("=== 项目架构最佳实践示例 ===");
            
            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();
                
                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheet;
                worksheet.Name = "架构设计";
                
                // 创建架构设计说明
                worksheet.Range["A1"].Value = "Excel自动化项目架构设计";
                worksheet.Range["A1"].Font.Size = 16;
                worksheet.Range["A1"].Font.Bold = true;
                worksheet.Range["A1"].Interior.Color = Color.Navy;
                worksheet.Range["A1"].Font.Color = Color.White;
                worksheet.Range["A1:G1"].Merge();
                
                // 分层架构说明
                worksheet.Range["A2"].Value = "分层架构";
                worksheet.Range["A2"].Font.Bold = true;
                worksheet.Range["A2"].Interior.Color = Color.LightBlue;
                worksheet.Range["A2:G2"].Merge();
                
                worksheet.Range["A3"].Value = "层级";
                worksheet.Range["B3"].Value = "职责";
                worksheet.Range["C3"].Value = "包含内容";
                worksheet.Range["D3"].Value = "技术要点";
                
                // 设置表头格式
                var headerRange = worksheet.Range["A3:D3"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 架构层级
                object[,] architectureLayers = {
                    {"表现层", "用户交互", "Web界面、控制台界面", "用户输入验证"},
                    {"应用层", "业务协调", "服务协调、工作流控制", "事务管理"},
                    {"业务层", "核心逻辑", "业务规则、数据处理", "领域模型"},
                    {"数据层", "数据访问", "Excel操作、数据持久化", "数据映射"},
                    {"基础设施层", "基础支持", "日志、配置、工具", "通用功能"}
                };
                
                var layerRange = worksheet.Range["A4:D8"];
                layerRange.Value = architectureLayers;
                layerRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 依赖注入配置
                worksheet.Range["A10"].Value = "依赖注入配置";
                worksheet.Range["A10"].Font.Bold = true;
                worksheet.Range["A10"].Interior.Color = Color.LightGreen;
                worksheet.Range["A10:G10"].Merge();
                
                worksheet.Range["A11"].Value = "服务类型";
                worksheet.Range["B11"].Value = "接口";
                worksheet.Range["C11"].Value = "实现类";
                worksheet.Range["D11"].Value = "生命周期";
                
                var diHeaderRange = worksheet.Range["A11:D11"];
                diHeaderRange.Font.Bold = true;
                diHeaderRange.Interior.Color = Color.LightGray;
                diHeaderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                object[,] diServices = {
                    {"核心服务", "IExcelApplication", "ExcelApplication", "Scoped"},
                    {"业务服务", "IReportGenerator", "ReportGenerator", "Scoped"},
                    {"工具服务", "IPerformanceMonitor", "PerformanceMonitor", "Singleton"},
                    {"日志服务", "ILogger", "FileLogger", "Singleton"}
                };
                
                var diRange = worksheet.Range["A12:D15"];
                diRange.Value = diServices;
                diRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 架构设计原则
                worksheet.Range["A17"].Value = "架构设计原则";
                worksheet.Range["A17"].Font.Bold = true;
                worksheet.Range["A17"].Interior.Color = Color.LightYellow;
                worksheet.Range["A17:G17"].Merge();
                
                worksheet.Range["A18"].Value = "1. 单一职责原则 - 每个类只负责一个功能";
                worksheet.Range["A19"].Value = "2. 开闭原则 - 对扩展开放，对修改关闭";
                worksheet.Range["A20"].Value = "3. 依赖倒置原则 - 依赖抽象而非具体实现";
                worksheet.Range["A21"].Value = "4. 接口隔离原则 - 使用专门的接口而非通用接口";
                worksheet.Range["A22"].Value = "5. 迪米特法则 - 降低类之间的耦合度";
                
                // 自动调整列宽
                worksheet.Columns.AutoFit();
                
                // 保存架构设计文件
                string architectureFileName = $"LayeredArchitecture_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(architectureFileName);
                
                Console.WriteLine($"✓ 成功演示项目架构最佳实践: {architectureFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 项目架构最佳实践出错: {ex.Message}");
            }
            
            Console.WriteLine();
        }

        /// <summary>
        /// 性能优化最佳实践示例
        /// 演示如何优化Excel自动化性能
        /// </summary>
        static void PerformanceOptimizationExample()
        {
            Console.WriteLine("=== 性能优化最佳实践示例 ===");
            
            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();
                
                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheet;
                worksheet.Name = "性能优化";
                
                // 创建性能优化说明
                worksheet.Range["A1"].Value = "Excel自动化性能优化最佳实践";
                worksheet.Range["A1"].Font.Size = 16;
                worksheet.Range["A1"].Font.Bold = true;
                worksheet.Range["A1"].Interior.Color = Color.DarkGreen;
                worksheet.Range["A1"].Font.Color = Color.White;
                worksheet.Range["A1:G1"].Merge();
                
                // 批量操作优化
                worksheet.Range["A2"].Value = "批量操作优化";
                worksheet.Range["A2"].Font.Bold = true;
                worksheet.Range["A2"].Interior.Color = Color.LightBlue;
                worksheet.Range["A2:G2"].Merge();
                
                worksheet.Range["A3"].Value = "优化技术";
                worksheet.Range["B3"].Value = "应用场景";
                worksheet.Range["C3"].Value = "性能提升";
                worksheet.Range["D3"].Value = "实现要点";
                
                // 设置表头格式
                var headerRange = worksheet.Range["A3:D3"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 优化技术
                object[,] optimizationTechniques = {
                    {"数组批量操作", "大量数据读写", "90%以上", "使用二维数组一次读写"},
                    {"禁用屏幕更新", "复杂计算处理", "50-80%", "设置ScreenUpdating=false"},
                    {"手动计算模式", "公式密集型", "70-90%", "设置Calculation=xlManual"},
                    {"对象缓存", "重复对象访问", "30-50%", "缓存常用对象引用"},
                    {"延迟释放", "批量COM操作", "20-40%", "操作完成后再释放对象"}
                };
                
                var techniqueRange = worksheet.Range["A4:D8"];
                techniqueRange.Value = optimizationTechniques;
                techniqueRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 内存管理优化
                worksheet.Range["A10"].Value = "内存管理优化";
                worksheet.Range["A10"].Font.Bold = true;
                worksheet.Range["A10"].Interior.Color = Color.LightGreen;
                worksheet.Range["A10:G10"].Merge();
                
                worksheet.Range["A11"].Value = "优化策略";
                worksheet.Range["B11"].Value = "实现方式";
                worksheet.Range["C11"].Value = "效果";
                
                var memoryHeaderRange = worksheet.Range["A11:C11"];
                memoryHeaderRange.Font.Bold = true;
                memoryHeaderRange.Interior.Color = Color.LightGray;
                memoryHeaderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                object[,] memoryOptimizations = {
                    {"及时释放COM对象", "using语句或显式Dispose", "避免内存泄漏"},
                    {"避免循环引用", "合理设计对象关系", "减少内存占用"},
                    {"批量处理数据", "分批处理大量数据", "控制内存峰值"},
                    {"使用值类型", "优先使用基本类型", "减少托管堆压力"}
                };
                
                var memoryRange = worksheet.Range["A12:C15"];
                memoryRange.Value = memoryOptimizations;
                memoryRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 性能监控
                worksheet.Range["A17"].Value = "性能监控";
                worksheet.Range["A17"].Font.Bold = true;
                worksheet.Range["A17"].Interior.Color = Color.LightYellow;
                worksheet.Range["A17:G17"].Merge();
                
                worksheet.Range["A18"].Value = "监控指标";
                worksheet.Range["B18"].Value = "监控方式";
                worksheet.Range["C18"].Value = "优化建议";
                
                var monitorHeaderRange = worksheet.Range["A18:C18"];
                monitorHeaderRange.Font.Bold = true;
                monitorHeaderRange.Interior.Color = Color.LightGray;
                monitorHeaderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                object[,] monitorMetrics = {
                    {"执行时间", "Stopwatch计时", "超过阈值时优化"},
                    {"内存使用", "PerformanceCounter", "持续增长时检查"},
                    {"CPU占用", "进程监控", "过高时分析热点"},
                    {"COM调用次数", "计数器统计", "过多时考虑批量操作"}
                };
                
                var monitorRange = worksheet.Range["A19:C22"];
                monitorRange.Value = monitorMetrics;
                monitorRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 自动调整列宽
                worksheet.Columns.AutoFit();
                
                // 保存性能优化文件
                string performanceFileName = $"PerformanceOptimization_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(performanceFileName);
                
                Console.WriteLine($"✓ 成功演示性能优化最佳实践: {performanceFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 性能优化最佳实践出错: {ex.Message}");
            }
            
            Console.WriteLine();
        }

        /// <summary>
        /// 错误处理最佳实践示例
        /// 演示如何正确处理Excel自动化中的错误
        /// </summary>
        static void ErrorHandlingExample()
        {
            Console.WriteLine("=== 错误处理最佳实践示例 ===");
            
            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();
                
                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheet;
                worksheet.Name = "错误处理";
                
                // 创建错误处理说明
                worksheet.Range["A1"].Value = "Excel自动化错误处理最佳实践";
                worksheet.Range["A1"].Font.Size = 16;
                worksheet.Range["A1"].Font.Bold = true;
                worksheet.Range["A1"].Interior.Color = Color.DarkRed;
                worksheet.Range["A1"].Font.Color = Color.White;
                worksheet.Range["A1:G1"].Merge();
                
                // 异常处理策略
                worksheet.Range["A2"].Value = "异常处理策略";
                worksheet.Range["A2"].Font.Bold = true;
                worksheet.Range["A2"].Interior.Color = Color.LightBlue;
                worksheet.Range["A2:G2"].Merge();
                
                worksheet.Range["A3"].Value = "异常类型";
                worksheet.Range["B3"].Value = "处理方式";
                worksheet.Range["C3"].Value = "恢复策略";
                worksheet.Range["D3"].Value = "日志记录";
                
                // 设置表头格式
                var headerRange = worksheet.Range["A3:D3"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 异常处理方式
                object[,] exceptionHandling = {
                    {"COMException", "重试机制", "重新创建对象", "详细错误信息"},
                    {"FileNotFoundException", "检查路径", "提供默认值", "文件路径信息"},
                    {"UnauthorizedAccessException", "权限检查", "提示用户", "操作权限信息"},
                    {"OutOfMemoryException", "资源释放", "重启服务", "内存使用情况"},
                    {"InvalidOperationException", "状态检查", "恢复状态", "操作上下文"}
                };
                
                var exceptionRange = worksheet.Range["A4:D8"];
                exceptionRange.Value = exceptionHandling;
                exceptionRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 资源清理策略
                worksheet.Range["A10"].Value = "资源清理策略";
                worksheet.Range["A10"].Font.Bold = true;
                worksheet.Range["A10"].Interior.Color = Color.LightGreen;
                worksheet.Range["A10:G10"].Merge();
                
                worksheet.Range["A11"].Value = "清理对象";
                worksheet.Range["B11"].Value = "清理方式";
                worksheet.Range["C11"].Value = "时机";
                worksheet.Range["D11"].Value = "注意事项";
                
                var cleanupHeaderRange = worksheet.Range["A11:D11"];
                cleanupHeaderRange.Font.Bold = true;
                cleanupHeaderRange.Interior.Color = Color.LightGray;
                cleanupHeaderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                object[,] cleanupStrategies = {
                    {"Excel应用程序", "Dispose方法", "使用完成后", "确保所有工作簿已关闭"},
                    {"工作簿", "Close方法", "操作完成后", "保存必要数据"},
                    {"工作表", "释放引用", "不再使用时", "避免循环引用"},
                    {"图表对象", "显式释放", "图表操作后", "防止内存泄漏"}
                };
                
                var cleanupRange = worksheet.Range["A12:D15"];
                cleanupRange.Value = cleanupStrategies;
                cleanupRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 日志记录规范
                worksheet.Range["A17"].Value = "日志记录规范";
                worksheet.Range["A17"].Font.Bold = true;
                worksheet.Range["A17"].Interior.Color = Color.LightYellow;
                worksheet.Range["A17:G17"].Merge();
                
                worksheet.Range["A18"].Value = "日志级别";
                worksheet.Range["B18"].Value = "记录内容";
                worksheet.Range["C18"].Value = "输出目标";
                worksheet.Range["D18"].Value = "保留期限";
                
                var logHeaderRange = worksheet.Range["A18:D18"];
                logHeaderRange.Font.Bold = true;
                logHeaderRange.Interior.Color = Color.LightGray;
                logHeaderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                object[,] logStandards = {
                    {"Error", "异常详细信息", "文件+控制台", "永久保留"},
                    {"Warn", "潜在问题", "文件", "1年"},
                    {"Info", "关键操作", "文件", "6个月"},
                    {"Debug", "调试信息", "文件", "1个月"}
                };
                
                var logRange = worksheet.Range["A19:D22"];
                logRange.Value = logStandards;
                logRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 自动调整列宽
                worksheet.Columns.AutoFit();
                
                // 保存错误处理文件
                string errorHandlingFileName = $"ErrorHandling_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(errorHandlingFileName);
                
                Console.WriteLine($"✓ 成功演示错误处理最佳实践: {errorHandlingFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 错误处理最佳实践出错: {ex.Message}");
            }
            
            Console.WriteLine();
        }

        /// <summary>
        /// 资源管理最佳实践示例
        /// 演示如何正确管理Excel自动化中的资源
        /// </summary>
        static void ResourceManagementExample()
        {
            Console.WriteLine("=== 资源管理最佳实践示例 ===");
            
            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();
                
                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheet;
                worksheet.Name = "资源管理";
                
                // 创建资源管理说明
                worksheet.Range["A1"].Value = "Excel自动化资源管理最佳实践";
                worksheet.Range["A1"].Font.Size = 16;
                worksheet.Range["A1"].Font.Bold = true;
                worksheet.Range["A1"].Interior.Color = Color.Purple;
                worksheet.Range["A1"].Font.Color = Color.White;
                worksheet.Range["A1:G1"].Merge();
                
                // COM对象管理
                worksheet.Range["A2"].Value = "COM对象管理";
                worksheet.Range["A2"].Font.Bold = true;
                worksheet.Range["A2"].Interior.Color = Color.LightBlue;
                worksheet.Range["A2:G2"].Merge();
                
                worksheet.Range["A3"].Value = "对象类型";
                worksheet.Range["B3"].Value = "创建方式";
                worksheet.Range["C3"].Value = "释放方式";
                worksheet.Range["D3"].Value = "注意事项";
                
                // 设置表头格式
                var headerRange = worksheet.Range["A3:D3"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // COM对象管理方式
                object[,] comObjectManagement = {
                    {"Excel应用程序", "ExcelFactory.Create()", "using语句", "确保DisplayAlerts=false"},
                    {"工作簿", "OpenWorkbook方法", "Close方法", "保存前检查路径权限"},
                    {"工作表", "Worksheets属性", "释放引用", "避免循环引用"},
                    {"图表", "Shapes.AddChart", "显式释放", "操作完成后立即释放"}
                };
                
                var comRange = worksheet.Range["A4:D7"];
                comRange.Value = comObjectManagement;
                comRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 内存管理策略
                worksheet.Range["A9"].Value = "内存管理策略";
                worksheet.Range["A9"].Font.Bold = true;
                worksheet.Range["A9"].Interior.Color = Color.LightGreen;
                worksheet.Range["A9:G9"].Merge();
                
                worksheet.Range["A10"].Value = "策略类型";
                worksheet.Range["B10"].Value = "实现方式";
                worksheet.Range["C10"].Value = "适用场景";
                worksheet.Range["D10"].Value = "效果";
                
                var memoryHeaderRange = worksheet.Range["A10:D10"];
                memoryHeaderRange.Font.Bold = true;
                memoryHeaderRange.Interior.Color = Color.LightGray;
                memoryHeaderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                object[,] memoryStrategies = {
                    {"对象池", "预创建对象复用", "频繁创建相同对象", "减少GC压力"},
                    {"延迟加载", "按需创建对象", "对象使用频率不均", "降低内存占用"},
                    {"及时释放", "操作后立即释放", "短期使用对象", "避免内存泄漏"},
                    {"批量处理", "集中处理数据", "大量数据操作", "提高处理效率"}
                };
                
                var memoryRange = worksheet.Range["A11:D14"];
                memoryRange.Value = memoryStrategies;
                memoryRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 文件资源管理
                worksheet.Range["A16"].Value = "文件资源管理";
                worksheet.Range["A16"].Font.Bold = true;
                worksheet.Range["A16"].Interior.Color = Color.LightYellow;
                worksheet.Range["A16:G16"].Merge();
                
                worksheet.Range["A17"].Value = "资源类型";
                worksheet.Range["B17"].Value = "管理方式";
                worksheet.Range["C17"].Value = "安全措施";
                worksheet.Range["D17"].Value = "清理策略";
                
                var fileHeaderRange = worksheet.Range["A17:D17"];
                fileHeaderRange.Font.Bold = true;
                fileHeaderRange.Interior.Color = Color.LightGray;
                fileHeaderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                object[,] fileManagement = {
                    {"临时文件", "Path.GetTempPath", "权限控制", "操作后删除"},
                    {"模板文件", "嵌入资源", "只读访问", "无需清理"},
                    {"输出文件", "指定目录", "备份机制", "定期清理"},
                    {"日志文件", "专用目录", "轮转策略", "按时间清理"}
                };
                
                var fileRange = worksheet.Range["A18:D21"];
                fileRange.Value = fileManagement;
                fileRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 自动调整列宽
                worksheet.Columns.AutoFit();
                
                // 保存资源管理文件
                string resourceManagementFileName = $"ResourceManagement_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(resourceManagementFileName);
                
                Console.WriteLine($"✓ 成功演示资源管理最佳实践: {resourceManagementFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 资源管理最佳实践出错: {ex.Message}");
            }
            
            Console.WriteLine();
        }

        /// <summary>
        /// 安全开发最佳实践示例
        /// 演示如何在Excel自动化开发中保证安全性
        /// </summary>
        static void SecurityBestPracticesExample()
        {
            Console.WriteLine("=== 安全开发最佳实践示例 ===");
            
            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();
                
                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheet;
                worksheet.Name = "安全开发";
                
                // 创建安全开发说明
                worksheet.Range["A1"].Value = "Excel自动化安全开发最佳实践";
                worksheet.Range["A1"].Font.Size = 16;
                worksheet.Range["A1"].Font.Bold = true;
                worksheet.Range["A1"].Interior.Color = Color.Orange;
                worksheet.Range["A1"].Font.Color = Color.White;
                worksheet.Range["A1:G1"].Merge();
                
                // 输入验证
                worksheet.Range["A2"].Value = "输入验证";
                worksheet.Range["A2"].Font.Bold = true;
                worksheet.Range["A2"].Interior.Color = Color.LightBlue;
                worksheet.Range["A2:G2"].Merge();
                
                worksheet.Range["A3"].Value = "验证类型";
                worksheet.Range["B3"].Value = "验证方法";
                worksheet.Range["C3"].Value = "验证规则";
                worksheet.Range["D3"].Value = "错误处理";
                
                // 设置表头格式
                var headerRange = worksheet.Range["A3:D3"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 输入验证方法
                object[,] inputValidation = {
                    {"文件路径", "Path.IsPathRooted", "合法路径格式", "返回错误信息"},
                    {"文件名", "白名单检查", "允许字符范围", "拒绝非法字符"},
                    {"数据范围", "数值检查", "最大最小值", "截断或提示"},
                    {"宏代码", "关键字过滤", "禁用危险函数", "拒绝执行"}
                };
                
                var validationRange = worksheet.Range["A4:D7"];
                validationRange.Value = inputValidation;
                validationRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 权限控制
                worksheet.Range["A9"].Value = "权限控制";
                worksheet.Range["A9"].Font.Bold = true;
                worksheet.Range["A9"].Interior.Color = Color.LightGreen;
                worksheet.Range["A9:G9"].Merge();
                
                worksheet.Range["A10"].Value = "控制类型";
                worksheet.Range["B10"].Value = "实现方式";
                worksheet.Range["C10"].Value = "控制粒度";
                worksheet.Range["D10"].Value = "审计要求";
                
                var permissionHeaderRange = worksheet.Range["A10:D10"];
                permissionHeaderRange.Font.Bold = true;
                permissionHeaderRange.Interior.Color = Color.LightGray;
                permissionHeaderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                object[,] permissionControl = {
                    {"文件访问", "文件系统权限", "目录级", "记录访问日志"},
                    {"功能使用", "角色权限", "功能级", "验证用户身份"},
                    {"数据操作", "数据权限", "记录级", "检查数据归属"},
                    {"宏执行", "安全设置", "应用级", "严格限制执行"}
                };
                
                var permissionRange = worksheet.Range["A11:D14"];
                permissionRange.Value = permissionControl;
                permissionRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 安全配置
                worksheet.Range["A16"].Value = "安全配置";
                worksheet.Range["A16"].Font.Bold = true;
                worksheet.Range["A16"].Interior.Color = Color.LightYellow;
                worksheet.Range["A16:G16"].Merge();
                
                worksheet.Range["A17"].Value = "配置项";
                worksheet.Range["B17"].Value = "安全设置";
                worksheet.Range["C17"].Value = "默认值";
                worksheet.Range["D17"].Value = "推荐值";
                
                var configHeaderRange = worksheet.Range["A17:D17"];
                configHeaderRange.Font.Bold = true;
                configHeaderRange.Interior.Color = Color.LightGray;
                configHeaderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                object[,] securityConfig = {
                    {"宏安全级别", "SecurityLevel", "低", "高"},
                    {"程序化访问", "AutomationSecurity", "启用", "提示用户"},
                    {"信任访问", "TrustAccess", "允许", "禁止"},
                    {"加载项", "AddIns", "自动加载", "手动加载"}
                };
                
                var configRange = worksheet.Range["A18:D21"];
                configRange.Value = securityConfig;
                configRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 自动调整列宽
                worksheet.Columns.AutoFit();
                
                // 保存安全开发文件
                string securityFileName = $"SecurityBestPractices_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(securityFileName);
                
                Console.WriteLine($"✓ 成功演示安全开发最佳实践: {securityFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 安全开发最佳实践出错: {ex.Message}");
            }
            
            Console.WriteLine();
        }

        /// <summary>
        /// 代码组织最佳实践示例
        /// 演示如何组织Excel自动化代码
        /// </summary>
        static void CodeOrganizationExample()
        {
            Console.WriteLine("=== 代码组织最佳实践示例 ===");
            
            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();
                
                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheet;
                worksheet.Name = "代码组织";
                
                // 创建代码组织说明
                worksheet.Range["A1"].Value = "Excel自动化代码组织最佳实践";
                worksheet.Range["A1"].Font.Size = 16;
                worksheet.Range["A1"].Font.Bold = true;
                worksheet.Range["A1"].Interior.Color = Color.Brown;
                worksheet.Range["A1"].Font.Color = Color.White;
                worksheet.Range["A1:G1"].Merge();
                
                // 命名规范
                worksheet.Range["A2"].Value = "命名规范";
                worksheet.Range["A2"].Font.Bold = true;
                worksheet.Range["A2"].Interior.Color = Color.LightBlue;
                worksheet.Range["A2:G2"].Merge();
                
                worksheet.Range["A3"].Value = "元素类型";
                worksheet.Range["B3"].Value = "命名规则";
                worksheet.Range["C3"].Value = "示例";
                worksheet.Range["D3"].Value = "说明";
                
                // 设置表头格式
                var headerRange = worksheet.Range["A3:D3"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 命名规范
                object[,] namingConventions = {
                    {"类名", "PascalCase", "ExcelReportGenerator", "名词或名词短语"},
                    {"方法名", "PascalCase", "GenerateReport", "动词或动词短语"},
                    {"变量名", "camelCase", "workbook", "描述性名称"},
                    {"常量名", "PascalCase", "MaxRows", "全大写也可接受"},
                    {"接口名", "I+PascalCase", "IExcelWorksheet", "I前缀表示接口"}
                };
                
                var namingRange = worksheet.Range["A4:D8"];
                namingRange.Value = namingConventions;
                namingRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 代码结构
                worksheet.Range["A10"].Value = "代码结构";
                worksheet.Range["A10"].Font.Bold = true;
                worksheet.Range["A10"].Interior.Color = Color.LightGreen;
                worksheet.Range["A10:G10"].Merge();
                
                worksheet.Range["A11"].Value = "结构元素";
                worksheet.Range["B11"].Value = "组织方式";
                worksheet.Range["C11"].Value = "示例";
                worksheet.Range["D11"].Value = "最佳实践";
                
                var structureHeaderRange = worksheet.Range["A11:D11"];
                structureHeaderRange.Font.Bold = true;
                structureHeaderRange.Interior.Color = Color.LightGray;
                structureHeaderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                object[,] codeStructure = {
                    {"命名空间", "功能分组", "MyApp.Excel.Reports", "按业务领域划分"},
                    {"类组织", "职责单一", "ChartService, DataService", "一个类一个职责"},
                    {"方法组织", "功能相关", "#region 图表操作", "使用region分组"},
                    {"属性封装", "访问控制", "private IExcelApp _app", "私有字段封装"},
                    {"注释规范", "XML文档", "/// <summary>", "详细描述功能"}
                };
                
                var structureRange = worksheet.Range["A12:D16"];
                structureRange.Value = codeStructure;
                structureRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 代码复用
                worksheet.Range["A18"].Value = "代码复用";
                worksheet.Range["A18"].Font.Bold = true;
                worksheet.Range["A18"].Interior.Color = Color.LightYellow;
                worksheet.Range["A18:G18"].Merge();
                
                worksheet.Range["A19"].Value = "复用方式";
                worksheet.Range["B19"].Value = "实现机制";
                worksheet.Range["C19"].Value = "适用场景";
                worksheet.Range["D19"].Value = "优势";
                
                var reuseHeaderRange = worksheet.Range["A19:D19"];
                reuseHeaderRange.Font.Bold = true;
                reuseHeaderRange.Interior.Color = Color.LightGray;
                reuseHeaderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                object[,] codeReuse = {
                    {"继承", "基类派生", "相似功能类", "代码共享"},
                    {"组合", "对象包含", "复杂功能封装", "灵活扩展"},
                    {"工具类", "静态方法", "通用功能", "方便调用"},
                    {"扩展方法", "this参数", "接口功能增强", "语法简洁"},
                    {"模板方法", "抽象基类", "算法骨架", "控制反转"}
                };
                
                var reuseRange = worksheet.Range["A20:D24"];
                reuseRange.Value = codeReuse;
                reuseRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 自动调整列宽
                worksheet.Columns.AutoFit();
                
                // 保存代码组织文件
                string codeOrganizationFileName = $"CodeOrganization_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(codeOrganizationFileName);
                
                Console.WriteLine($"✓ 成功演示代码组织最佳实践: {codeOrganizationFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 代码组织最佳实践出错: {ex.Message}");
            }
            
            Console.WriteLine();
        }
    }
}
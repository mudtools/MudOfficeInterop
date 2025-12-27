# 第15篇：宏与自动化脚本详解

## 引言：Excel自动化的"智能机器人"

在Excel自动化开发中，如果说基础操作是"手动驾驶"，那么宏与自动化脚本就是"自动驾驶"！它们就像是给Excel装上了"智能机器人"，能够自动执行复杂的业务逻辑，实现真正的无人值守操作。

想象一下这样的场景：每天凌晨2点，当整个办公室都沉浸在宁静中时，你的Excel自动化系统却正在"辛勤工作"——自动下载最新的销售数据、执行复杂的计算分析、生成精美的报表、甚至自动发送邮件给相关领导。这一切都不需要人工干预，就像有一个不知疲倦的机器人24小时为你服务！

MudTools.OfficeInterop.Excel项目就像是专业的"智能机器人制造工厂"，它提供了完整的宏执行和自动化脚本支持。从简单的Excel 4.0宏到复杂的VBA脚本，从基础的自动化任务到高级的业务流程，每一个功能都能让你的Excel自动化达到新的高度。

本篇将带你探索宏与自动化脚本的奥秘，学习如何通过代码创建智能、高效、可靠的企业级自动化解决方案。准备好让你的Excel系统拥有"自主思考"和"自动执行"的能力了吗？

## 宏执行基础

### 宏执行接口定义

项目通过`IExcelApplication`接口提供了完整的宏执行功能：

```csharp
/// <summary>
/// 执行Excel 4.0宏函数
/// 对应 Application.ExecuteExcel4Macro 方法
/// </summary>
/// <param name="macro">宏函数</param>
/// <returns>执行结果</returns>
object ExecuteExcel4Macro(string macro);
```

### 宏执行管理器

```csharp
using System;
using System.Collections.Generic;
using MudTools.OfficeInterop.Excel.CoreComponents.Core;
using MudTools.OfficeInterop.Excel.Enums.Application;

namespace MudTools.OfficeInterop.Excel.AdvancedFeatures.Macro
{
    /// <summary>
    /// 宏执行管理器
    /// 提供Excel宏的创建、执行和管理功能
    /// </summary>
    public class MacroExecutionManager
    {
        private readonly IExcelApplication _application;
        private readonly Dictionary<string, MacroInfo> _macros;
        
        public MacroExecutionManager(IExcelApplication? Application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _macros = new Dictionary<string, MacroInfo>();
        }
        
        /// <summary>
        /// 执行Excel 4.0宏
        /// </summary>
        public object ExecuteExcel4Macro(string macroCode)
        {
            if (string.IsNullOrWhiteSpace(macroCode))
                throw new ArgumentException("宏代码不能为空", nameof(macroCode));
            
            try
            {
                // 记录宏执行统计
                _application.PerformanceStats.MacroExecutions++;
                
                // 执行Excel 4.0宏
                return _application.ExecuteExcel4Macro(macroCode);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"执行Excel 4.0宏失败: {ex.Message}", ex);
            }
        }
        
        /// <summary>
        /// 执行VBA宏
        /// </summary>
        public object ExecuteVbaMacro(string macroName, params object[] args)
        {
            if (string.IsNullOrWhiteSpace(macroName))
                throw new ArgumentException("宏名称不能为空", nameof(macroName));
            
            try
            {
                // 记录宏执行统计
                _application.PerformanceStats.MacroExecutions++;
                
                // 执行VBA宏
                return _application.Run(macroName, args);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"执行VBA宏失败: {ex.Message}", ex);
            }
        }
        
        /// <summary>
        /// 注册宏信息
        /// </summary>
        public void RegisterMacro(string name, string description, MacroType type)
        {
            _macros[name] = new MacroInfo(name, description, type);
        }
        
        /// <summary>
        /// 获取所有已注册的宏
        /// </summary>
        public IEnumerable<MacroInfo> GetRegisteredMacros()
        {
            return _macros.Values;
        }
        
        /// <summary>
        /// 宏信息类
        /// </summary>
        public class MacroInfo
        {
            public string Name { get; }
            public string Description { get; }
            public MacroType Type { get; }
            
            public MacroInfo(string name, string description, MacroType type)
            {
                Name = name;
                Description = description;
                Type = type;
            }
        }
    }
    
    /// <summary>
    /// 宏类型枚举
    /// </summary>
    public enum MacroType
    {
        Excel4Macro,
        VbaMacro,
        CustomScript
    }
}
```

## Excel 4.0宏应用

### 基础Excel 4.0宏函数

```csharp
/// <summary>
/// Excel 4.0宏函数管理器
/// 提供常用的Excel 4.0宏函数封装
/// </summary>
public class Excel4MacroManager
{
    private readonly MacroExecutionManager _macroManager;
    
    public Excel4MacroManager(MacroExecutionManager macroManager)
    {
        _macroManager = macroManager;
    }
    
    /// <summary>
    /// 设置单元格值
    /// </summary>
    public void SetCellValue(string cellAddress, object value)
    {
        string macroCode = $"SET.VALUE({cellAddress}, {value})";
        _macroManager.ExecuteExcel4Macro(macroCode);
    }
    
    /// <summary>
    /// 获取单元格值
    /// </summary>
    public object GetCellValue(string cellAddress)
    {
        string macroCode = $"GET.CELL(1, {cellAddress})";
        return _macroManager.ExecuteExcel4Macro(macroCode);
    }
    
    /// <summary>
    /// 选择单元格区域
    /// </summary>
    public void SelectRange(string rangeAddress)
    {
        string macroCode = $"SELECT({rangeAddress})";
        _macroManager.ExecuteExcel4Macro(macroCode);
    }
    
    /// <summary>
    /// 激活工作表
    /// </summary>
    public void ActivateSheet(string sheetName)
    {
        string macroCode = $"ACTIVATE(\"{sheetName}\")";
        _macroManager.ExecuteExcel4Macro(macroCode);
    }
    
    /// <summary>
    /// 创建工作表
    /// </summary>
    public void CreateSheet(string sheetName)
    {
        string macroCode = $"WORKBOOK.INSERT(1)";
        _macroManager.ExecuteExcel4Macro(macroCode);
        
        // 重命名新工作表
        string renameMacro = $"WORKBOOK.NAME(\"{sheetName}\")";
        _macroManager.ExecuteExcel4Macro(renameMacro);
    }
    
    /// <summary>
    /// 删除工作表
    /// </summary>
    public void DeleteSheet(string sheetName)
    {
        string macroCode = $"WORKBOOK.DELETE(\"{sheetName}\")";
        _macroManager.ExecuteExcel4Macro(macroCode);
    }
    
    /// <summary>
    /// 保存工作簿
    /// </summary>
    public void SaveWorkbook(string filePath = null)
    {
        string macroCode = string.IsNullOrEmpty(filePath) 
            ? "FILE.SAVE()" 
            : $"FILE.SAVE(\"{filePath}\")";
        
        _macroManager.ExecuteExcel4Macro(macroCode);
    }
    
    /// <summary>
    /// 打开工作簿
    /// </summary>
    public void OpenWorkbook(string filePath)
    {
        string macroCode = $"FILE.OPEN(\"{filePath}\")";
        _macroManager.ExecuteExcel4Macro(macroCode);
    }
    
    /// <summary>
    /// 关闭工作簿
    /// </summary>
    public void CloseWorkbook(string workbookName = null)
    {
        string macroCode = string.IsNullOrEmpty(workbookName)
            ? "FILE.CLOSE()"
            : $"FILE.CLOSE(\"{workbookName}\")";
        
        _macroManager.ExecuteExcel4Macro(macroCode);
    }
    
    /// <summary>
    /// 执行计算
    /// </summary>
    public void Calculate()
    {
        string macroCode = "CALCULATE.NOW()";
        _macroManager.ExecuteExcel4Macro(macroCode);
    }
    
    /// <summary>
    /// 设置打印区域
    /// </summary>
    public void SetPrintArea(string rangeAddress)
    {
        string macroCode = $"SET.PRINT.AREA({rangeAddress})";
        _macroManager.ExecuteExcel4Macro(macroCode);
    }
    
    /// <summary>
    /// 打印工作表
    /// </summary>
    public void PrintSheet()
    {
        string macroCode = "PRINT()";
        _macroManager.ExecuteExcel4Macro(macroCode);
    }
}
```

### 高级Excel 4.0宏应用

```csharp
/// <summary>
/// 高级Excel 4.0宏应用管理器
/// 提供复杂的Excel 4.0宏功能封装
/// </summary>
public class AdvancedExcel4MacroManager
{
    private readonly Excel4MacroManager _baseManager;
    
    public AdvancedExcel4MacroManager(Excel4MacroManager baseManager)
    {
        _baseManager = baseManager;
    }
    
    /// <summary>
    /// 创建数据透视表
    /// </summary>
    public void CreatePivotTable(string sourceRange, string destinationCell, 
        string rowFields, string columnFields, string dataFields)
    {
        string macroCode = $"""
            CREATE.OBJECT(2,1,1,1,1,1,1)
            PIVOT.ADD.DATA({sourceRange})
            PIVOT.TABLE.WIZARD(1,{destinationCell},{rowFields},{columnFields},{dataFields})
            """;
        
        _baseManager.ExecuteMacro(macroCode);
    }
    
    /// <summary>
    /// 创建图表
    /// </summary>
    public void CreateChart(string dataRange, string chartType, string title)
    {
        string macroCode = $"""
            CREATE.OBJECT(3,1,1,1,1,1,1)
            CHART.WIZARD({dataRange},{chartType},1,1,1,1,1,1,1,1,1,1)
            CHART.TITLE(\"{title}\")
            """;
        
        _baseManager.ExecuteMacro(macroCode);
    }
    
    /// <summary>
    /// 数据排序
    /// </summary>
    public void SortData(string dataRange, string sortBy, bool ascending = true)
    {
        string order = ascending ? "1" : "2";
        string macroCode = $"SORT({dataRange},{sortBy},{order})";
        
        _baseManager.ExecuteMacro(macroCode);
    }
    
    /// <summary>
    /// 数据筛选
    /// </summary>
    public void FilterData(string dataRange, string criteriaRange)
    {
        string macroCode = $"DATA.FILTER({dataRange},{criteriaRange})";
        
        _baseManager.ExecuteMacro(macroCode);
    }
    
    /// <summary>
    /// 数据验证
    /// </summary>
    public void SetDataValidation(string cellRange, string validationType, string criteria)
    {
        string macroCode = $"DATA.VALIDATION({cellRange},{validationType},{criteria})";
        
        _baseManager.ExecuteMacro(macroCode);
    }
    
    /// <summary>
    /// 条件格式设置
    /// </summary>
    public void SetConditionalFormatting(string cellRange, string condition, string format)
    {
        string macroCode = $"FORMAT.CONDITIONAL({cellRange},{condition},{format})";
        
        _baseManager.ExecuteMacro(macroCode);
    }
    
    /// <summary>
    /// 宏录制功能模拟
    /// </summary>
    public void StartMacroRecording(string macroName)
    {
        string macroCode = $"RECORD(\"{macroName}\")";
        _baseManager.ExecuteMacro(macroCode);
    }
    
    /// <summary>
    /// 停止宏录制
    /// </summary>
    public void StopMacroRecording()
    {
        string macroCode = "STOP.RECORD()";
        _baseManager.ExecuteMacro(macroCode);
    }
    
    /// <summary>
    /// 执行录制的宏
    /// </summary>
    public void ExecuteRecordedMacro(string macroName)
    {
        string macroCode = $"RUN(\"{macroName}\")";
        _baseManager.ExecuteMacro(macroCode);
    }
}
```

## VBA宏集成

### VBA宏执行管理器

```csharp
/// <summary>
/// VBA宏执行管理器
/// 提供VBA宏的创建、执行和管理功能
/// </summary>
public class VbaMacroManager
{
    private readonly MacroExecutionManager _macroManager;
    private readonly Dictionary<string, VbaMacroInfo> _vbaMacros;
    
    public VbaMacroManager(MacroExecutionManager macroManager)
    {
        _macroManager = macroManager;
        _vbaMacros = new Dictionary<string, VbaMacroInfo>();
    }
    
    /// <summary>
    /// 执行VBA宏
    /// </summary>
    public object ExecuteVbaMacro(string macroName, params object[] args)
    {
        if (string.IsNullOrWhiteSpace(macroName))
            throw new ArgumentException("VBA宏名称不能为空", nameof(macroName));
        
        try
        {
            return _macroManager.ExecuteVbaMacro(macroName, args);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"执行VBA宏'{macroName}'失败: {ex.Message}", ex);
        }
    }
    
    /// <summary>
    /// 注册VBA宏
    /// </summary>
    public void RegisterVbaMacro(string name, string description, string moduleName, 
        string procedureName, VbaMacroType type)
    {
        _vbaMacros[name] = new VbaMacroInfo(name, description, moduleName, procedureName, type);
    }
    
    /// <summary>
    /// 获取VBA宏信息
    /// </summary>
    public VbaMacroInfo GetVbaMacroInfo(string name)
    {
        return _vbaMacros.TryGetValue(name, out var info) ? info : null;
    }
    
    /// <summary>
    /// 执行带参数的VBA函数
    /// </summary>
    public object ExecuteVbaFunction(string functionName, params object[] parameters)
    {
        // 构建完整的函数调用
        string fullFunctionName = $"ThisWorkbook.{functionName}";
        return ExecuteVbaMacro(fullFunctionName, parameters);
    }
    
    /// <summary>
    /// 执行工作表级别的VBA宏
    /// </summary>
    public object ExecuteWorksheetMacro(string sheetName, string macroName, params object[] args)
    {
        string fullMacroName = $"'{sheetName}'!{macroName}";
        return ExecuteVbaMacro(fullMacroName, args);
    }
    
    /// <summary>
    /// 执行模块级别的VBA宏
    /// </summary>
    public object ExecuteModuleMacro(string moduleName, string macroName, params object[] args)
    {
        string fullMacroName = $"{moduleName}.{macroName}";
        return ExecuteVbaMacro(fullMacroName, args);
    }
    
    /// <summary>
    /// VBA宏信息类
    /// </summary>
    public class VbaMacroInfo
    {
        public string Name { get; }
        public string Description { get; }
        public string ModuleName { get; }
        public string ProcedureName { get; }
        public VbaMacroType Type { get; }
        
        public VbaMacroInfo(string name, string description, string moduleName, 
            string procedureName, VbaMacroType type)
        {
            Name = name;
            Description = description;
            ModuleName = moduleName;
            ProcedureName = procedureName;
            Type = type;
        }
        
        /// <summary>
        /// 获取完整的宏名称
        /// </summary>
        public string FullName => $"{ModuleName}.{ProcedureName}";
    }
    
    /// <summary>
    /// VBA宏类型枚举
    /// </summary>
    public enum VbaMacroType
    {
        SubProcedure,    // Sub过程
        Function,        // Function函数
        EventHandler    // 事件处理程序
    }
}
```

### 常用VBA宏封装

```csharp
/// <summary>
/// 常用VBA宏封装管理器
/// 提供常用VBA功能的封装
/// </summary>
public class CommonVbaMacroManager
{
    private readonly VbaMacroManager _vbaManager;
    
    public CommonVbaMacroManager(VbaMacroManager vbaManager)
    {
        _vbaManager = vbaManager;
    }
    
    /// <summary>
    /// 显示消息框
    /// </summary>
    public void ShowMessageBox(string message, string title = "提示")
    {
        _vbaManager.ExecuteVbaMacro("ShowMessageBox", message, title);
    }
    
    /// <summary>
    /// 输入框
    /// </summary>
    public string ShowInputBox(string prompt, string title = "输入")
    {
        var result = _vbaManager.ExecuteVbaMacro("ShowInputBox", prompt, title);
        return result?.ToString() ?? string.Empty;
    }
    
    /// <summary>
    /// 文件选择对话框
    /// </summary>
    public string ShowFileDialog(string initialPath = "", string filter = "所有文件 (*.*)|*.*")
    {
        var result = _vbaManager.ExecuteVbaMacro("ShowFileDialog", initialPath, filter);
        return result?.ToString() ?? string.Empty;
    }
    
    /// <summary>
    /// 文件夹选择对话框
    /// </summary>
    public string ShowFolderDialog(string initialPath = "")
    {
        var result = _vbaManager.ExecuteVbaMacro("ShowFolderDialog", initialPath);
        return result?.ToString() ?? string.Empty;
    }
    
    /// <summary>
    /// 保存工作簿
    /// </summary>
    public void SaveWorkbookAs(string filePath)
    {
        _vbaManager.ExecuteVbaMacro("SaveWorkbookAs", filePath);
    }
    
    /// <summary>
    /// 导出工作表为PDF
    /// </summary>
    public void ExportToPdf(string sheetName, string filePath)
    {
        _vbaManager.ExecuteVbaMacro("ExportToPdf", sheetName, filePath);
    }
    
    /// <summary>
    /// 发送邮件
    /// </summary>
    public void SendEmail(string to, string subject, string body, string attachmentPath = "")
    {
        _vbaManager.ExecuteVbaMacro("SendEmail", to, subject, body, attachmentPath);
    }
    
    /// <summary>
    /// 数据验证
    /// </summary>
    public bool ValidateData(string dataRange, string validationRule)
    {
        var result = _vbaManager.ExecuteVbaMacro("ValidateData", dataRange, validationRule);
        return Convert.ToBoolean(result);
    }
    
    /// <summary>
    /// 数据清理
    /// </summary>
    public void CleanData(string dataRange)
    {
        _vbaManager.ExecuteVbaMacro("CleanData", dataRange);
    }
    
    /// <summary>
    /// 数据转换
    /// </summary>
    public void TransformData(string sourceRange, string destinationRange, string transformationRule)
    {
        _vbaManager.ExecuteVbaMacro("TransformData", sourceRange, destinationRange, transformationRule);
    }
    
    /// <summary>
    /// 生成报告
    /// </summary>
    public void GenerateReport(string dataRange, string templatePath, string outputPath)
    {
        _vbaManager.ExecuteVbaMacro("GenerateReport", dataRange, templatePath, outputPath);
    }
    
    /// <summary>
    /// 备份数据
    /// </summary>
    public void BackupData(string dataRange, string backupPath)
    {
        _vbaManager.ExecuteVbaMacro("BackupData", dataRange, backupPath);
    }
    
    /// <summary>
    /// 恢复数据
    /// </summary>
    public void RestoreData(string backupPath, string destinationRange)
    {
        _vbaManager.ExecuteVbaMacro("RestoreData", backupPath, destinationRange);
    }
}
```

## 自动化脚本系统

### 脚本执行引擎

```csharp
/// <summary>
/// 自动化脚本执行引擎
/// 提供脚本的解析、执行和管理功能
/// </summary>
public class AutomationScriptEngine
{
    private readonly MacroExecutionManager _macroManager;
    private readonly VbaMacroManager _vbaManager;
    private readonly Excel4MacroManager _excel4Manager;
    private readonly Dictionary<string, AutomationScript> _scripts;
    
    public AutomationScriptEngine(MacroExecutionManager macroManager, 
        VbaMacroManager vbaManager, Excel4MacroManager excel4Manager)
    {
        _macroManager = macroManager;
        _vbaManager = vbaManager;
        _excel4Manager = excel4Manager;
        _scripts = new Dictionary<string, AutomationScript>();
    }
    
    /// <summary>
    /// 注册自动化脚本
    /// </summary>
    public void RegisterScript(string name, AutomationScript script)
    {
        _scripts[name] = script;
    }
    
    /// <summary>
    /// 执行自动化脚本
    /// </summary>
    public ScriptExecutionResult ExecuteScript(string scriptName, Dictionary<string, object> parameters = null)
    {
        if (!_scripts.TryGetValue(scriptName, out var script))
            throw new ArgumentException($"脚本'{scriptName}'未注册");
        
        var result = new ScriptExecutionResult(scriptName);
        
        try
        {
            result.StartTime = DateTime.Now;
            
            // 执行脚本步骤
            foreach (var step in script.Steps)
            {
                var stepResult = ExecuteScriptStep(step, parameters);
                result.StepResults.Add(stepResult);
                
                if (!stepResult.Success)
                {
                    result.Success = false;
                    result.ErrorMessage = stepResult.ErrorMessage;
                    break;
                }
            }
            
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
    /// 执行脚本步骤
    /// </summary>
    private ScriptStepResult ExecuteScriptStep(ScriptStep step, Dictionary<string, object> parameters)
    {
        var stepResult = new ScriptStepResult(step.Name);
        
        try
        {
            switch (step.Type)
            {
                case ScriptStepType.Excel4Macro:
                    stepResult.Result = _excel4Manager.ExecuteMacro(step.Code);
                    break;
                    
                case ScriptStepType.VbaMacro:
                    var macroParameters = PrepareMacroParameters(step.Parameters, parameters);
                    stepResult.Result = _vbaManager.ExecuteVbaMacro(step.Code, macroParameters);
                    break;
                    
                case ScriptStepType.Condition:
                    stepResult.Result = EvaluateCondition(step.Code, parameters);
                    break;
                    
                case ScriptStepType.Loop:
                    stepResult.Result = ExecuteLoop(step, parameters);
                    break;
                    
                default:
                    throw new NotSupportedException($"不支持的脚本步骤类型: {step.Type}");
            }
            
            stepResult.Success = true;
        }
        catch (Exception ex)
        {
            stepResult.Success = false;
            stepResult.ErrorMessage = ex.Message;
            stepResult.Exception = ex;
        }
        
        return stepResult;
    }
    
    /// <summary>
    /// 准备宏参数
    /// </summary>
    private object[] PrepareMacroParameters(Dictionary<string, string> stepParameters, 
        Dictionary<string, object> scriptParameters)
    {
        var parameters = new List<object>();
        
        foreach (var param in stepParameters)
        {
            if (scriptParameters != null && scriptParameters.ContainsKey(param.Value))
            {
                parameters.Add(scriptParameters[param.Value]);
            }
            else
            {
                parameters.Add(param.Value);
            }
        }
        
        return parameters.ToArray();
    }
    
    /// <summary>
    /// 评估条件
    /// </summary>
    private bool EvaluateCondition(string condition, Dictionary<string, object> parameters)
    {
        // 简单的条件评估逻辑
        // 实际应用中可以使用更复杂的表达式解析器
        return true; // 简化实现
    }
    
    /// <summary>
    /// 执行循环
    /// </summary>
    private object ExecuteLoop(ScriptStep step, Dictionary<string, object> parameters)
    {
        // 简化实现
        return null;
    }
    
    /// <summary>
    /// 获取所有已注册的脚本
    /// </summary>
    public IEnumerable<AutomationScript> GetRegisteredScripts()
    {
        return _scripts.Values;
    }
    
    /// <summary>
    /// 获取脚本信息
    /// </summary>
    public AutomationScript GetScriptInfo(string name)
    {
        return _scripts.TryGetValue(name, out var script) ? script : null;
    }
}

/// <summary>
/// 自动化脚本类
/// </summary>
public class AutomationScript
{
    public string Name { get; set; }
    public string Description { get; set; }
    public List<ScriptStep> Steps { get; set; }
    public Dictionary<string, string> Parameters { get; set; }
    
    public AutomationScript()
    {
        Steps = new List<ScriptStep>();
        Parameters = new Dictionary<string, string>();
    }
}

/// <summary>
/// 脚本步骤类
/// </summary>
public class ScriptStep
{
    public string Name { get; set; }
    public ScriptStepType Type { get; set; }
    public string Code { get; set; }
    public Dictionary<string, string> Parameters { get; set; }
    
    public ScriptStep()
    {
        Parameters = new Dictionary<string, string>();
    }
}

/// <summary>
/// 脚本步骤类型枚举
/// </summary>
public enum ScriptStepType
{
    Excel4Macro,    // Excel 4.0宏
    VbaMacro,       // VBA宏
    Condition,      // 条件判断
    Loop            // 循环控制
}

/// <summary>
/// 脚本执行结果类
/// </summary>
public class ScriptExecutionResult
{
    public string ScriptName { get; }
    public bool Success { get; set; }
    public DateTime StartTime { get; set; }
    public DateTime EndTime { get; set; }
    public TimeSpan Duration { get; set; }
    public string ErrorMessage { get; set; }
    public Exception Exception { get; set; }
    public List<ScriptStepResult> StepResults { get; set; }
    
    public ScriptExecutionResult(string scriptName)
    {
        ScriptName = scriptName;
        StepResults = new List<ScriptStepResult>();
    }
}

/// <summary>
/// 脚本步骤执行结果类
/// </summary>
public class ScriptStepResult
{
    public string StepName { get; }
    public bool Success { get; set; }
    public object Result { get; set; }
    public string ErrorMessage { get; set; }
    public Exception Exception { get; set; }
    
    public ScriptStepResult(string stepName)
    {
        StepName = stepName;
    }
}
```

### 预定义自动化脚本

```csharp
/// <summary>
/// 预定义自动化脚本管理器
/// 提供常用的自动化脚本模板
/// </summary>
public class PredefinedScriptManager
{
    private readonly AutomationScriptEngine _scriptEngine;
    
    public PredefinedScriptManager(AutomationScriptEngine scriptEngine)
    {
        _scriptEngine = scriptEngine;
        RegisterPredefinedScripts();
    }
    
    /// <summary>
    /// 注册预定义脚本
    /// </summary>
    private void RegisterPredefinedScripts()
    {
        // 数据导入脚本
        RegisterDataImportScripts();
        
        // 数据处理脚本
        RegisterDataProcessingScripts();
        
        // 报告生成脚本
        RegisterReportGenerationScripts();
        
        // 系统维护脚本
        RegisterSystemMaintenanceScripts();
    }
    
    /// <summary>
    /// 注册数据导入脚本
    /// </summary>
    private void RegisterDataImportScripts()
    {
        var csvImportScript = new AutomationScript
        {
            Name = "CSV数据导入",
            Description = "从CSV文件导入数据到Excel工作表"
        };
        
        csvImportScript.Steps.Add(new ScriptStep
        {
            Name = "选择CSV文件",
            Type = ScriptStepType.VbaMacro,
            Code = "ShowFileDialog"
        });
        
        csvImportScript.Steps.Add(new ScriptStep
        {
            Name = "导入数据",
            Type = ScriptStepType.VbaMacro,
            Code = "ImportCsvData"
        });
        
        _scriptEngine.RegisterScript("csv_import", csvImportScript);
    }
    
    /// <summary>
    /// 注册数据处理脚本
    /// </summary>
    private void RegisterDataProcessingScripts()
    {
        var dataCleanupScript = new AutomationScript
        {
            Name = "数据清理",
            Description = "自动清理和标准化数据"
        };
        
        dataCleanupScript.Steps.Add(new ScriptStep
        {
            Name = "验证数据",
            Type = ScriptStepType.VbaMacro,
            Code = "ValidateData"
        });
        
        dataCleanupScript.Steps.Add(new ScriptStep
        {
            Name = "清理数据",
            Type = ScriptStepType.VbaMacro,
            Code = "CleanData"
        });
        
        _scriptEngine.RegisterScript("data_cleanup", dataCleanupScript);
    }
    
    /// <summary>
    /// 注册报告生成脚本
    /// </summary>
    private void RegisterReportGenerationScripts()
    {
        var salesReportScript = new AutomationScript
        {
            Name = "销售报告生成",
            Description = "自动生成销售分析报告"
        };
        
        salesReportScript.Steps.Add(new ScriptStep
        {
            Name = "汇总数据",
            Type = ScriptStepType.VbaMacro,
            Code = "SummarizeSalesData"
        });
        
        salesReportScript.Steps.Add(new ScriptStep
        {
            Name = "生成图表",
            Type = ScriptStepType.VbaMacro,
            Code = "CreateSalesCharts"
        });
        
        salesReportScript.Steps.Add(new ScriptStep
        {
            Name = "导出报告",
            Type = ScriptStepType.VbaMacro,
            Code = "ExportSalesReport"
        });
        
        _scriptEngine.RegisterScript("sales_report", salesReportScript);
    }
    
    /// <summary>
    /// 注册系统维护脚本
    /// </summary>
    private void RegisterSystemMaintenanceScripts()
    {
        var backupScript = new AutomationScript
        {
            Name = "数据备份",
            Description = "自动备份重要数据"
        };
        
        backupScript.Steps.Add(new ScriptStep
        {
            Name = "选择备份目录",
            Type = ScriptStepType.VbaMacro,
            Code = "ShowFolderDialog"
        });
        
        backupScript.Steps.Add(new ScriptStep
        {
            Name = "执行备份",
            Type = ScriptStepType.VbaMacro,
            Code = "BackupData"
        });
        
        _scriptEngine.RegisterScript("data_backup", backupScript);
    }
    
    /// <summary>
    /// 执行预定义脚本
    /// </summary>
    public ScriptExecutionResult ExecutePredefinedScript(string scriptName, 
        Dictionary<string, object> parameters = null)
    {
        return _scriptEngine.ExecuteScript(scriptName, parameters);
    }
    
    /// <summary>
    /// 获取所有预定义脚本
    /// </summary>
    public IEnumerable<AutomationScript> GetPredefinedScripts()
    {
        return _scriptEngine.GetRegisteredScripts()
            .Where(s => s.Name.StartsWith("预定义"));
    }
}
```

## 性能优化和错误处理

### 宏执行性能优化

```csharp
/// <summary>
/// 宏执行性能优化器
/// 提供宏执行的性能监控和优化功能
/// </summary>
public class MacroPerformanceOptimizer
{
    private readonly IExcelApplication _application;
    private readonly Dictionary<string, PerformanceMetrics> _performanceMetrics;
    
    public MacroPerformanceOptimizer(IExcelApplication? Application)
    {
        _application = application;
        _performanceMetrics = new Dictionary<string, PerformanceMetrics>();
    }
    
    /// <summary>
    /// 执行带性能监控的宏
    /// </summary>
    public object ExecuteMacroWithMonitoring(string macroName, Func<object> macroAction)
    {
        var metrics = new PerformanceMetrics(macroName);
        
        try
        {
            metrics.StartTimer();
            
            // 禁用屏幕更新提高性能
            _application.ScreenUpdating = false;
            
            // 禁用事件处理
            _application.EnableEvents = false;
            
            // 执行宏
            var result = macroAction();
            
            metrics.StopTimer();
            metrics.Success = true;
            
            return result;
        }
        catch (Exception ex)
        {
            metrics.StopTimer();
            metrics.Success = false;
            metrics.ErrorMessage = ex.Message;
            
            throw;
        }
        finally
        {
            // 恢复设置
            _application.ScreenUpdating = true;
            _application.EnableEvents = true;
            
            // 记录性能指标
            _performanceMetrics[macroName] = metrics;
        }
    }
    
    /// <summary>
    /// 批量执行宏优化
    /// </summary>
    public void ExecuteMacrosInBatch(IEnumerable<MacroExecutionTask> tasks)
    {
        try
        {
            // 批量优化设置
            _application.ScreenUpdating = false;
            _application.EnableEvents = false;
            _application.Calculation = CalculationMode.Manual;
            
            foreach (var task in tasks)
            {
                var metrics = new PerformanceMetrics(task.MacroName);
                metrics.StartTimer();
                
                try
                {
                    task.Action();
                    metrics.Success = true;
                }
                catch (Exception ex)
                {
                    metrics.Success = false;
                    metrics.ErrorMessage = ex.Message;
                    
                    if (task.StopOnError)
                        throw;
                }
                finally
                {
                    metrics.StopTimer();
                    _performanceMetrics[task.MacroName] = metrics;
                }
            }
        }
        finally
        {
            // 恢复设置
            _application.ScreenUpdating = true;
            _application.EnableEvents = true;
            _application.Calculation = CalculationMode.Automatic;
            
            // 强制重新计算
            _application.Calculate();
        }
    }
    
    /// <summary>
    /// 获取性能报告
    /// </summary>
    public PerformanceReport GetPerformanceReport()
    {
        var report = new PerformanceReport();
        
        foreach (var metrics in _performanceMetrics.Values)
        {
            report.AddMetric(metrics);
        }
        
        return report;
    }
    
    /// <summary>
    /// 宏执行任务类
    /// </summary>
    public class MacroExecutionTask
    {
        public string MacroName { get; set; }
        public Action Action { get; set; }
        public bool StopOnError { get; set; } = true;
    }
    
    /// <summary>
    /// 性能指标类
    /// </summary>
    public class PerformanceMetrics
    {
        public string MacroName { get; }
        public DateTime StartTime { get; private set; }
        public DateTime EndTime { get; private set; }
        public TimeSpan Duration { get; private set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        
        public PerformanceMetrics(string macroName)
        {
            MacroName = macroName;
        }
        
        public void StartTimer()
        {
            StartTime = DateTime.Now;
        }
        
        public void StopTimer()
        {
            EndTime = DateTime.Now;
            Duration = EndTime - StartTime;
        }
    }
    
    /// <summary>
    /// 性能报告类
    /// </summary>
    public class PerformanceReport
    {
        public List<PerformanceMetrics> Metrics { get; set; }
        public TimeSpan TotalDuration => TimeSpan.FromTicks(Metrics.Sum(m => m.Duration.Ticks));
        public int TotalExecutions => Metrics.Count;
        public int SuccessfulExecutions => Metrics.Count(m => m.Success);
        public int FailedExecutions => Metrics.Count(m => !m.Success);
        
        public PerformanceReport()
        {
            Metrics = new List<PerformanceMetrics>();
        }
        
        public void AddMetric(PerformanceMetrics metric)
        {
            Metrics.Add(metric);
        }
    }
}
```

### 错误处理和日志记录

```csharp
/// <summary>
/// 宏错误处理器
/// 提供宏执行的错误处理和日志记录功能
/// </summary>
public class MacroErrorHandler
{
    private readonly IExcelApplication _application;
    private readonly List<MacroError> _errors;
    
    public MacroErrorHandler(IExcelApplication? Application)
    {
        _application = application;
        _errors = new List<MacroError>();
    }
    
    /// <summary>
    /// 安全执行宏
    /// </summary>
    public MacroExecutionResult ExecuteMacroSafely(string macroName, Func<object> macroAction)
    {
        var result = new MacroExecutionResult(macroName);
        
        try
        {
            result.StartTime = DateTime.Now;
            result.Result = macroAction();
            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
            result.Exception = ex;
            
            // 记录错误
            var error = new MacroError(macroName, ex);
            _errors.Add(error);
            
            // 显示错误信息
            ShowErrorMessage(error);
        }
        finally
        {
            result.EndTime = DateTime.Now;
            result.Duration = result.EndTime - result.StartTime;
        }
        
        return result;
    }
    
    /// <summary>
    /// 显示错误信息
    /// </summary>
    private void ShowErrorMessage(MacroError error)
    {
        try
        {
            // 使用VBA显示错误对话框
            _application.Run("ShowErrorDialog", error.Message, error.MacroName);
        }
        catch
        {
            // 如果VBA不可用，使用基础错误处理
            Console.WriteLine($"宏执行错误: {error.Message}");
        }
    }
    
    /// <summary>
    /// 获取错误报告
    /// </summary>
    public ErrorReport GetErrorReport()
    {
        var report = new ErrorReport();
        
        foreach (var error in _errors)
        {
            report.AddError(error);
        }
        
        return report;
    }
    
    /// <summary>
    /// 清除错误记录
    /// </summary>
    public void ClearErrors()
    {
        _errors.Clear();
    }
    
    /// <summary>
    /// 宏错误类
    /// </summary>
    public class MacroError
    {
        public string MacroName { get; }
        public string Message { get; }
        public DateTime Timestamp { get; }
        public Exception Exception { get; }
        
        public MacroError(string macroName, Exception exception)
        {
            MacroName = macroName;
            Message = exception.Message;
            Timestamp = DateTime.Now;
            Exception = exception;
        }
    }
    
    /// <summary>
    /// 宏执行结果类
    /// </summary>
    public class MacroExecutionResult
    {
        public string MacroName { get; }
        public bool Success { get; set; }
        public object Result { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan Duration { get; set; }
        public string ErrorMessage { get; set; }
        public Exception Exception { get; set; }
        
        public MacroExecutionResult(string macroName)
        {
            MacroName = macroName;
        }
    }
    
    /// <summary>
    /// 错误报告类
    /// </summary>
    public class ErrorReport
    {
        public List<MacroError> Errors { get; set; }
        public int TotalErrors => Errors.Count;
        public DateTime ReportTime { get; }
        
        public ErrorReport()
        {
            Errors = new List<MacroError>();
            ReportTime = DateTime.Now;
        }
        
        public void AddError(MacroError error)
        {
            Errors.Add(error);
        }
        
        /// <summary>
        /// 获取按宏名称分组的错误统计
        /// </summary>
        public Dictionary<string, int> GetErrorStatistics()
        {
            return Errors
                .GroupBy(e => e.MacroName)
                .ToDictionary(g => g.Key, g => g.Count());
        }
    }
}
```

## 实际应用案例

### 销售数据分析自动化系统

```csharp
/// <summary>
/// 销售数据分析自动化系统
/// 完整的销售数据自动化处理解决方案
/// </summary>
public class SalesDataAutomationSystem
{
    private readonly MacroExecutionManager _macroManager;
    private readonly VbaMacroManager _vbaManager;
    private readonly AutomationScriptEngine _scriptEngine;
    private readonly MacroPerformanceOptimizer _performanceOptimizer;
    private readonly MacroErrorHandler _errorHandler;
    
    public SalesDataAutomationSystem(IExcelApplication? Application)
    {
        _macroManager = new MacroExecutionManager(application);
        _vbaManager = new VbaMacroManager(_macroManager);
        var excel4Manager = new Excel4MacroManager(_macroManager);
        _scriptEngine = new AutomationScriptEngine(_macroManager, _vbaManager, excel4Manager);
        _performanceOptimizer = new MacroPerformanceOptimizer(application);
        _errorHandler = new MacroErrorHandler(application);
        
        InitializeSystem();
    }
    
    /// <summary>
    /// 初始化系统
    /// </summary>
    private void InitializeSystem()
    {
        // 注册销售相关的VBA宏
        RegisterSalesMacros();
        
        // 注册销售自动化脚本
        RegisterSalesScripts();
    }
    
    /// <summary>
    /// 注册销售相关的VBA宏
    /// </summary>
    private void RegisterSalesMacros()
    {
        _vbaManager.RegisterVbaMacro(
            "ImportSalesData", 
            "导入销售数据", 
            "SalesModule", 
            "ImportSalesData", 
            VbaMacroManager.VbaMacroType.Function);
            
        _vbaManager.RegisterVbaMacro(
            "AnalyzeSalesTrends", 
            "分析销售趋势", 
            "SalesModule", 
            "AnalyzeSalesTrends", 
            VbaMacroManager.VbaMacroType.Function);
            
        _vbaManager.RegisterVbaMacro(
            "GenerateSalesReport", 
            "生成销售报告", 
            "SalesModule", 
            "GenerateSalesReport", 
            VbaMacroManager.VbaMacroType.Function);
    }
    
    /// <summary>
    /// 注册销售自动化脚本
    /// </summary>
    private void RegisterSalesScripts()
    {
        var dailySalesScript = new AutomationScript
        {
            Name = "每日销售数据处理",
            Description = "自动处理每日销售数据，生成分析报告"
        };
        
        dailySalesScript.Steps.Add(new ScriptStep
        {
            Name = "导入销售数据",
            Type = ScriptStepType.VbaMacro,
            Code = "ImportSalesData"
        });
        
        dailySalesScript.Steps.Add(new ScriptStep
        {
            Name = "数据验证和清理",
            Type = ScriptStepType.VbaMacro,
            Code = "ValidateAndCleanSalesData"
        });
        
        dailySalesScript.Steps.Add(new ScriptStep
        {
            Name = "销售趋势分析",
            Type = ScriptStepType.VbaMacro,
            Code = "AnalyzeSalesTrends"
        });
        
        dailySalesScript.Steps.Add(new ScriptStep
        {
            Name = "生成销售报告",
            Type = ScriptStepType.VbaMacro,
            Code = "GenerateSalesReport"
        });
        
        _scriptEngine.RegisterScript("daily_sales_processing", dailySalesScript);
    }
    
    /// <summary>
    /// 执行每日销售数据处理
    /// </summary>
    public ScriptExecutionResult ProcessDailySalesData(Dictionary<string, object> parameters = null)
    {
        return _errorHandler.ExecuteMacroSafely("每日销售数据处理", () =>
        {
            return _performanceOptimizer.ExecuteMacroWithMonitoring(
                "daily_sales_processing", 
                () => _scriptEngine.ExecuteScript("daily_sales_processing", parameters));
        }).Result as ScriptExecutionResult;
    }
    
    /// <summary>
    /// 获取系统性能报告
    /// </summary>
    public MacroPerformanceOptimizer.PerformanceReport GetPerformanceReport()
    {
        return _performanceOptimizer.GetPerformanceReport();
    }
    
    /// <summary>
    /// 获取错误报告
    /// </summary>
    public MacroErrorHandler.ErrorReport GetErrorReport()
    {
        return _errorHandler.GetErrorReport();
    }
    
    /// <summary>
    /// 清除错误记录
    /// </summary>
    public void ClearErrorRecords()
    {
        _errorHandler.ClearErrors();
    }
}
```

## 总结

本篇博文详细介绍了MudTools.OfficeInterop.Excel项目中宏与自动化脚本的完整实现方案，包括：

1. **宏执行基础**：Excel 4.0宏和VBA宏的完整执行框架
2. **高级功能**：复杂的宏应用场景和自动化脚本系统
3. **性能优化**：专业的性能监控和优化技术
4. **错误处理**：完善的错误处理和日志记录机制
5. **实际应用**：完整的销售数据分析自动化系统案例

通过这些技术，开发者可以构建高效、可靠的企业级Excel自动化解决方案，显著提升工作效率和数据处理的准确性。
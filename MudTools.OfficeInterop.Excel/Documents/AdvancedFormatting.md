# 第九篇：高级格式化技巧详解

## 引言：Excel自动化的"智能管家"

在Excel自动化开发中，如果说基础格式化是"化妆"，那么高级格式化就是"智能管家"！它不仅能让报表看起来更美观，更重要的是能够确保数据质量、增强交互体验、提升工作效率。

想象一下这样的场景：你创建了一个销售数据录入系统，但用户可能会输入错误的数据类型、超出范围的数值、或者不符合规范的文本。传统的手动检查不仅效率低下，而且容易遗漏错误。但通过高级格式化技术，你可以设置智能的数据验证规则，让系统自动检查每一个输入，确保数据的准确性和一致性。

更令人兴奋的是，高级格式化还包括超链接管理、注释添加、条件格式等强大功能。这些功能就像是给Excel装上了"智能大脑"，让它能够理解业务逻辑，自动完成复杂的格式化任务。

本篇将带你探索高级格式化的奥秘，学习如何通过代码创建智能、交互、专业的Excel报表。准备好让你的Excel自动化系统拥有"智能管家"的能力了吗？

## 数据验证规则设置

### 数据验证基础概念

数据验证是确保数据质量的关键技术，通过设置验证规则可以限制用户输入的内容，防止无效数据进入系统。

#### 验证类型枚举

```csharp
// 数据验证类型枚举
public enum XlDVType
{
    xlValidateInputOnly = 0,      // 仅验证输入
    xlValidateWholeNumber = 1,    // 整数验证
    xlValidateDecimal = 2,        // 小数验证
    xlValidateList = 3,           // 列表验证
    xlValidateDate = 4,           // 日期验证
    xlValidateTime = 5,           // 时间验证
    xlValidateTextLength = 6,     // 文本长度验证
    xlValidateCustom = 7          // 自定义公式验证
}

// 警告样式枚举
public enum XlDVAlertStyle
{
    xlValidAlertStop = 1,         // 停止样式（禁止输入）
    xlValidAlertWarning = 2,      // 警告样式（允许选择）
    xlValidAlertInformation = 3   // 信息样式（仅提示）
}
```

### 数据验证管理器实现

```csharp
using MudTools.OfficeInterop.Excel;
using System;
using System.Collections.Generic;

/// <summary>
/// 数据验证管理器
/// 提供完整的数据验证规则设置和管理功能
/// </summary>
public class DataValidationManager
{
    private readonly IExcelWorksheet _worksheet;
    
    public DataValidationManager(IExcelWorksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }
    
    /// <summary>
    /// 设置整数验证规则
    /// </summary>
    public void SetIntegerValidation(string rangeAddress, int minValue, int maxValue, 
                                    string inputTitle = "", string inputMessage = "",
                                    string errorTitle = "输入错误", string errorMessage = "请输入有效的整数")
    {
        var range = _worksheet.Range(rangeAddress);
        if (range == null) return;
        
        // 清除现有验证规则
        ClearValidation(rangeAddress);
        
        // 设置整数验证
        var validation = range.Validation;
        if (validation != null)
        {
            validation.Add(XlDVType.xlValidateWholeNumber, XlDVAlertStyle.xlValidAlertStop, 
                          minValue.ToString(), maxValue.ToString());
            
            // 设置输入提示
            if (!string.IsNullOrEmpty(inputTitle))
                validation.InputTitle = inputTitle;
            if (!string.IsNullOrEmpty(inputMessage))
                validation.InputMessage = inputMessage;
            
            // 设置错误提示
            validation.ErrorTitle = errorTitle;
            validation.ErrorMessage = errorMessage;
            validation.ShowError = true;
            validation.IgnoreBlank = true;
        }
    }
    
    /// <summary>
    /// 设置列表验证规则
    /// </summary>
    public void SetListValidation(string rangeAddress, List<string> items, 
                                string inputTitle = "", string inputMessage = "")
    {
        var range = _worksheet.Range(rangeAddress);
        if (range == null) return;
        
        ClearValidation(rangeAddress);
        
        var validation = range.Validation;
        if (validation != null)
        {
            // 创建列表公式（逗号分隔的列表）
            string listFormula = string.Join(",", items);
            
            validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, 
                          listFormula);
            
            validation.InCellDropdown = true; // 显示下拉箭头
            
            if (!string.IsNullOrEmpty(inputTitle))
                validation.InputTitle = inputTitle;
            if (!string.IsNullOrEmpty(inputMessage))
                validation.InputMessage = inputMessage;
        }
    }
    
    /// <summary>
    /// 设置日期验证规则
    /// </summary>
    public void SetDateValidation(string rangeAddress, DateTime startDate, DateTime endDate,
                                string errorMessage = "请输入有效的日期")
    {
        var range = _worksheet.Range(rangeAddress);
        if (range == null) return;
        
        ClearValidation(rangeAddress);
        
        var validation = range.Validation;
        if (validation != null)
        {
            validation.Add(XlDVType.xlValidateDate, XlDVAlertStyle.xlValidAlertStop,
                          startDate.ToOADate().ToString(), endDate.ToOADate().ToString());
            
            validation.ErrorMessage = errorMessage;
            validation.ShowError = true;
        }
    }
    
    /// <summary>
    /// 设置自定义公式验证
    /// </summary>
    public void SetCustomValidation(string rangeAddress, string formula, 
                                  string errorMessage = "数据验证失败")
    {
        var range = _worksheet.Range(rangeAddress);
        if (range == null) return;
        
        ClearValidation(rangeAddress);
        
        var validation = range.Validation;
        if (validation != null)
        {
            validation.Add(XlDVType.xlValidateCustom, XlDVAlertStyle.xlValidAlertStop, formula);
            
            validation.ErrorMessage = errorMessage;
            validation.ShowError = true;
        }
    }
    
    /// <summary>
    /// 清除验证规则
    /// </summary>
    public void ClearValidation(string rangeAddress)
    {
        var range = _worksheet.Range(rangeAddress);
        if (range == null) return;
        
        var validation = range.Validation;
        validation?.Delete();
    }
    
    /// <summary>
    /// 批量设置验证规则
    /// </summary>
    public void BatchSetValidations(Dictionary<string, ValidationRule> rules)
    {
        foreach (var rule in rules)
        {
            switch (rule.Type)
            {
                case ValidationType.Integer:
                    SetIntegerValidation(rule.RangeAddress, rule.MinValue, rule.MaxValue, 
                                      rule.InputTitle, rule.InputMessage);
                    break;
                case ValidationType.List:
                    SetListValidation(rule.RangeAddress, rule.ListItems, 
                                    rule.InputTitle, rule.InputMessage);
                    break;
                case ValidationType.Date:
                    SetDateValidation(rule.RangeAddress, rule.StartDate, rule.EndDate);
                    break;
                case ValidationType.Custom:
                    SetCustomValidation(rule.RangeAddress, rule.CustomFormula);
                    break;
            }
        }
    }
}

/// <summary>
/// 验证规则定义
/// </summary>
public class ValidationRule
{
    public string RangeAddress { get; set; } = string.Empty;
    public ValidationType Type { get; set; }
    public int MinValue { get; set; }
    public int MaxValue { get; set; }
    public List<string> ListItems { get; set; } = new List<string>();
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
    public string CustomFormula { get; set; } = string.Empty;
    public string InputTitle { get; set; } = string.Empty;
    public string InputMessage { get; set; } = string.Empty;
}

public enum ValidationType
{
    Integer,
    List,
    Date,
    Custom
}
```

### 数据验证应用场景

#### 员工信息录入系统

```csharp
public class EmployeeDataEntrySystem
{
    private readonly IExcelWorksheet _worksheet;
    private readonly DataValidationManager _validationManager;
    
    public EmployeeDataEntrySystem(IExcelWorksheet worksheet)
    {
        _worksheet = worksheet;
        _validationManager = new DataValidationManager(worksheet);
        
        SetupValidationRules();
    }
    
    private void SetupValidationRules()
    {
        // 员工编号验证（5位数字）
        _validationManager.SetCustomValidation("A2:A100", "=AND(LEN(A2)=5,ISNUMBER(A2))", 
                                              "员工编号必须是5位数字");
        
        // 部门选择验证
        var departments = new List<string> { "技术部", "销售部", "财务部", "人事部", "行政部" };
        _validationManager.SetListValidation("B2:B100", departments, "选择部门", "请从下拉列表中选择部门");
        
        // 入职日期验证（2010年至今）
        _validationManager.SetDateValidation("C2:C100", new DateTime(2010, 1, 1), DateTime.Today,
                                            "请输入2010年至今的有效日期");
        
        // 薪资验证（2000-50000）
        _validationManager.SetIntegerValidation("D2:D100", 2000, 50000, "薪资输入", 
                                              "请输入2000-50000之间的整数");
        
        // 邮箱格式验证
        _validationManager.SetCustomValidation("E2:E100", 
                                              "=AND(ISNUMBER(FIND("@",E2)),ISNUMBER(FIND(".",E2)))",
                                              "请输入有效的邮箱地址");
    }
    
    /// <summary>
    /// 验证所有数据
    /// </summary>
    public bool ValidateAllData()
    {
        bool isValid = true;
        
        // 这里可以实现更复杂的验证逻辑
        // 比如检查数据完整性、业务规则等
        
        return isValid;
    }
}
```

## 超链接管理

### 超链接基础操作

超链接是创建交互式文档的重要工具，可以链接到网页、文件、邮件地址或文档内部位置。

```csharp
using MudTools.OfficeInterop.Excel;
using System;
using System.Collections.Generic;

/// <summary>
/// 超链接管理器
/// 提供完整的超链接创建和管理功能
/// </summary>
public class HyperlinkManager
{
    private readonly IExcelWorksheet _worksheet;
    
    public HyperlinkManager(IExcelWorksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }
    
    /// <summary>
    /// 创建网页链接
    /// </summary>
    public void CreateWebLink(string cellAddress, string url, string displayText = "")
    {
        var range = _worksheet.Range(cellAddress);
        if (range == null) return;
        
        var hyperlinks = range.Hyperlinks;
        if (hyperlinks != null)
        {
            string text = string.IsNullOrEmpty(displayText) ? url : displayText;
            hyperlinks.Add(range, url, null, null, text);
            
            // 设置单元格格式（蓝色下划线）
            range.Font.Color = RGB(0, 0, 255); // 蓝色
            range.Font.Underline = true;
        }
    }
    
    /// <summary>
    /// 创建邮件链接
    /// </summary>
    public void CreateEmailLink(string cellAddress, string email, string subject = "", 
                               string displayText = "")
    {
        var range = _worksheet.Range(cellAddress);
        if (range == null) return;
        
        string mailtoUrl = $"mailto:{email}";
        if (!string.IsNullOrEmpty(subject))
        {
            mailtoUrl += $"?subject={Uri.EscapeDataString(subject)}";
        }
        
        string text = string.IsNullOrEmpty(displayText) ? email : displayText;
        
        var hyperlinks = range.Hyperlinks;
        hyperlinks?.Add(range, mailtoUrl, null, null, text);
        
        // 设置格式
        range.Font.Color = RGB(0, 0, 255);
        range.Font.Underline = true;
    }
    
    /// <summary>
    /// 创建文件链接
    /// </summary>
    public void CreateFileLink(string cellAddress, string filePath, string displayText = "")
    {
        var range = _worksheet.Range(cellAddress);
        if (range == null) return;
        
        string text = string.IsNullOrEmpty(displayText) ? System.IO.Path.GetFileName(filePath) : displayText;
        
        var hyperlinks = range.Hyperlinks;
        hyperlinks?.Add(range, filePath, null, null, text);
        
        range.Font.Color = RGB(0, 0, 255);
        range.Font.Underline = true;
    }
    
    /// <summary>
    /// 创建工作表内部链接
    /// </summary>
    public void CreateInternalLink(string cellAddress, string targetCell, string displayText = "")
    {
        var range = _worksheet.Range(cellAddress);
        if (range == null) return;
        
        string text = string.IsNullOrEmpty(displayText) ? targetCell : displayText;
        
        var hyperlinks = range.Hyperlinks;
        hyperlinks?.Add(range, "", targetCell, null, text);
        
        range.Font.Color = RGB(0, 0, 128); // 深蓝色
        range.Font.Underline = true;
    }
    
    /// <summary>
    /// 获取所有超链接
    /// </summary>
    public List<HyperlinkInfo> GetAllHyperlinks()
    {
        var hyperlinks = new List<HyperlinkInfo>();
        
        var worksheetHyperlinks = _worksheet.Hyperlinks;
        if (worksheetHyperlinks != null)
        {
            for (int i = 1; i <= worksheetHyperlinks.Count; i++)
            {
                var link = worksheetHyperlinks[i];
                if (link != null)
                {
                    hyperlinks.Add(new HyperlinkInfo
                    {
                        Address = link.Address,
                        SubAddress = link.SubAddress,
                        TextToDisplay = link.TextToDisplay,
                        ScreenTip = link.ScreenTip,
                        RangeAddress = link.Range?.Address
                    });
                }
            }
        }
        
        return hyperlinks;
    }
    
    /// <summary>
    /// 删除超链接
    /// </summary>
    public void DeleteHyperlink(string cellAddress)
    {
        var range = _worksheet.Range(cellAddress);
        if (range == null) return;
        
        var hyperlinks = range.Hyperlinks;
        hyperlinks?.Delete();
        
        // 恢复默认格式
        range.Font.Color = RGB(0, 0, 0); // 黑色
        range.Font.Underline = false;
    }
    
    /// <summary>
    /// 批量创建超链接
    /// </summary>
    public void BatchCreateHyperlinks(List<HyperlinkConfig> configs)
    {
        foreach (var config in configs)
        {
            switch (config.LinkType)
            {
                case HyperlinkType.Web:
                    CreateWebLink(config.CellAddress, config.Target, config.DisplayText);
                    break;
                case HyperlinkType.Email:
                    CreateEmailLink(config.CellAddress, config.Target, config.Subject, config.DisplayText);
                    break;
                case HyperlinkType.File:
                    CreateFileLink(config.CellAddress, config.Target, config.DisplayText);
                    break;
                case HyperlinkType.Internal:
                    CreateInternalLink(config.CellAddress, config.Target, config.DisplayText);
                    break;
            }
        }
    }
    
    // 辅助方法：RGB颜色转换
    private int RGB(int red, int green, int blue)
    {
        return (red & 0xFF) | ((green & 0xFF) << 8) | ((blue & 0xFF) << 16);
    }
}

/// <summary>
/// 超链接信息
/// </summary>
public class HyperlinkInfo
{
    public string Address { get; set; } = string.Empty;
    public string SubAddress { get; set; } = string.Empty;
    public string TextToDisplay { get; set; } = string.Empty;
    public string ScreenTip { get; set; } = string.Empty;
    public string RangeAddress { get; set; } = string.Empty;
}

/// <summary>
/// 超链接配置
/// </summary>
public class HyperlinkConfig
{
    public string CellAddress { get; set; } = string.Empty;
    public HyperlinkType LinkType { get; set; }
    public string Target { get; set; } = string.Empty;
    public string DisplayText { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
}

public enum HyperlinkType
{
    Web,
    Email,
    File,
    Internal
}
```

### 超链接应用场景

#### 项目文档管理系统

```csharp
public class ProjectDocumentManager
{
    private readonly IExcelWorksheet _worksheet;
    private readonly HyperlinkManager _hyperlinkManager;
    
    public ProjectDocumentManager(IExcelWorksheet worksheet)
    {
        _worksheet = worksheet;
        _hyperlinkManager = new HyperlinkManager(worksheet);
        
        SetupProjectLinks();
    }
    
    private void SetupProjectLinks()
    {
        // 项目主页链接
        _hyperlinkManager.CreateWebLink("A2", "https://company.com/projects/alpha", "项目主页");
        
        // 项目文档链接
        _hyperlinkManager.CreateFileLink("B2", @"\\server\projects\alpha\spec.docx", "需求文档");
        _hyperlinkManager.CreateFileLink("C2", @"\\server\projects\alpha\design.pptx", "设计文档");
        _hyperlinkManager.CreateFileLink("D2", @"\\server\projects\alpha\test.xlsx", "测试报告");
        
        // 相关资源链接
        _hyperlinkManager.CreateWebLink("E2", "https://docs.microsoft.com", "技术文档");
        _hyperlinkManager.CreateWebLink("F2", "https://stackoverflow.com", "技术问答");
        
        // 联系人邮件链接
        _hyperlinkManager.CreateEmailLink("G2", "project.manager@company.com", "项目Alpha相关问题", 
                                        "联系项目经理");
        _hyperlinkManager.CreateEmailLink("H2", "tech.lead@company.com", "技术问题咨询", 
                                        "联系技术负责人");
        
        // 内部导航链接
        _hyperlinkManager.CreateInternalLink("I2", "A50", "跳转到项目总结");
        _hyperlinkManager.CreateInternalLink("J2", "B100", "跳转到风险分析");
    }
    
    /// <summary>
    /// 添加新文档链接
    /// </summary>
    public void AddDocumentLink(string documentName, string filePath, string cellAddress)
    {
        _hyperlinkManager.CreateFileLink(cellAddress, filePath, documentName);
    }
    
    /// <summary>
    /// 验证所有链接的有效性
    /// </summary>
    public List<string> ValidateLinks()
    {
        var invalidLinks = new List<string>();
        var allLinks = _hyperlinkManager.GetAllHyperlinks();
        
        foreach (var link in allLinks)
        {
            if (!IsLinkValid(link))
            {
                invalidLinks.Add($"{link.RangeAddress}: {link.TextToDisplay}");
            }
        }
        
        return invalidLinks;
    }
    
    private bool IsLinkValid(HyperlinkInfo link)
    {
        // 实现链接有效性检查逻辑
        // 可以检查文件是否存在、网址是否可达等
        return true; // 简化实现
    }
}
```

## 注释管理

### 注释基础操作

注释是Excel中用于添加说明和备注的重要功能，可以增强文档的可读性和协作性。

```csharp
using MudTools.OfficeInterop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;

/// <summary>
/// 注释管理器
/// 提供完整的注释创建、编辑和管理功能
/// </summary>
public class CommentManager
{
    private readonly IExcelWorksheet _worksheet;
    
    public CommentManager(IExcelWorksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }
    
    /// <summary>
    /// 添加注释
    /// </summary>
    public void AddComment(string cellAddress, string commentText, string author = "系统")
    {
        var range = _worksheet.Range(cellAddress);
        if (range == null) return;
        
        var comment = range.AddComment(commentText);
        if (comment != null)
        {
            comment.Author = author;
            
            // 设置注释格式
            comment.Visible = false; // 默认隐藏
            comment.Shape.Fill.ForeColor.RGB = RGB(255, 255, 225); // 浅黄色背景
            comment.Shape.Line.ForeColor.RGB = RGB(0, 0, 0); // 黑色边框
        }
    }
    
    /// <summary>
    /// 编辑注释
    /// </summary>
    public void EditComment(string cellAddress, string newText, string author = "")
    {
        var range = _worksheet.Range(cellAddress);
        if (range == null) return;
        
        var comment = range.Comment;
        if (comment != null)
        {
            comment.Text(newText + $"\n编辑时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            
            if (!string.IsNullOrEmpty(author))
            {
                comment.Author = author;
            }
        }
    }
    
    /// <summary>
    /// 删除注释
    /// </summary>
    public void DeleteComment(string cellAddress)
    {
        var range = _worksheet.Range(cellAddress);
        if (range == null) return;
        
        var comment = range.Comment;
        comment?.Delete();
    }
    
    /// <summary>
    /// 显示/隐藏注释
    /// </summary>
    public void ToggleCommentVisibility(string cellAddress, bool visible)
    {
        var range = _worksheet.Range(cellAddress);
        if (range == null) return;
        
        var comment = range.Comment;
        if (comment != null)
        {
            comment.Visible = visible;
        }
    }
    
    /// <summary>
    /// 设置注释格式
    /// </summary>
    public void FormatComment(string cellAddress, CommentFormat format)
    {
        var range = _worksheet.Range(cellAddress);
        if (range == null) return;
        
        var comment = range.Comment;
        if (comment != null)
        {
            var shape = comment.Shape;
            
            // 背景颜色
            if (format.BackgroundColor.HasValue)
            {
                shape.Fill.ForeColor.RGB = format.BackgroundColor.Value;
            }
            
            // 边框颜色
            if (format.BorderColor.HasValue)
            {
                shape.Line.ForeColor.RGB = format.BorderColor.Value;
            }
            
            // 字体设置
            if (format.FontSize > 0)
            {
                shape.TextFrame.Characters().Font.Size = format.FontSize;
            }
            
            if (!string.IsNullOrEmpty(format.FontName))
            {
                shape.TextFrame.Characters().Font.Name = format.FontName;
            }
            
            // 自动调整大小
            shape.TextFrame.AutoSize = true;
        }
    }
    
    /// <summary>
    /// 获取所有注释
    /// </summary>
    public List<CommentInfo> GetAllComments()
    {
        var comments = new List<CommentInfo>();
        
        // 遍历所有包含注释的单元格
        var usedRange = _worksheet.UsedRange;
        if (usedRange != null)
        {
            foreach (var cell in usedRange.Cells)
            {
                var comment = cell.Comment;
                if (comment != null)
                {
                    comments.Add(new CommentInfo
                    {
                        CellAddress = cell.Address,
                        Text = comment.Text(),
                        Author = comment.Author,
                        Visible = comment.Visible
                    });
                }
            }
        }
        
        return comments;
    }
    
    /// <summary>
    /// 批量添加注释
    /// </summary>
    public void BatchAddComments(List<CommentConfig> configs)
    {
        foreach (var config in configs)
        {
            AddComment(config.CellAddress, config.Text, config.Author);
            
            if (config.Format != null)
            {
                FormatComment(config.CellAddress, config.Format);
            }
        }
    }
    
    /// <summary>
    /// 导出注释到文本文件
    /// </summary>
    public void ExportCommentsToFile(string filePath)
    {
        var comments = GetAllComments();
        
        using (var writer = new System.IO.StreamWriter(filePath))
        {
            writer.WriteLine("单元格地址\t作者\t注释内容");
            writer.WriteLine(new string('-', 80));
            
            foreach (var comment in comments)
            {
                writer.WriteLine($"{comment.CellAddress}\t{comment.Author}\t{comment.Text.Replace("\n", " ")}");
            }
        }
    }
    
    // 辅助方法：RGB颜色转换
    private int RGB(int red, int green, int blue)
    {
        return (red & 0xFF) | ((green & 0xFF) << 8) | ((blue & 0xFF) << 16);
    }
}

/// <summary>
/// 注释信息
/// </summary>
public class CommentInfo
{
    public string CellAddress { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public string Author { get; set; } = string.Empty;
    public bool Visible { get; set; }
}

/// <summary>
/// 注释配置
/// </summary>
public class CommentConfig
{
    public string CellAddress { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public string Author { get; set; } = "系统";
    public CommentFormat Format { get; set; } = new CommentFormat();
}

/// <summary>
/// 注释格式
/// </summary>
public class CommentFormat
{
    public int? BackgroundColor { get; set; } = RGB(255, 255, 225); // 浅黄色
    public int? BorderColor { get; set; } = RGB(0, 0, 0); // 黑色
    public float FontSize { get; set; } = 10;
    public string FontName { get; set; } = "宋体";
    
    private static int RGB(int red, int green, int blue)
    {
        return (red & 0xFF) | ((green & 0xFF) << 8) | ((blue & 0xFF) << 16);
    }
}
```

### 注释应用场景

#### 财务报表注释系统

```csharp
public class FinancialReportCommentSystem
{
    private readonly IExcelWorksheet _worksheet;
    private readonly CommentManager _commentManager;
    
    public FinancialReportCommentSystem(IExcelWorksheet worksheet)
    {
        _worksheet = worksheet;
        _commentManager = new CommentManager(worksheet);
        
        SetupFinancialComments();
    }
    
    private void SetupFinancialComments()
    {
        // 收入相关注释
        _commentManager.AddComment("B5", "本期收入增长主要来自新产品销售", "财务分析师");
        _commentManager.AddComment("B6", "服务收入保持稳定增长", "财务分析师");
        
        // 成本相关注释
        _commentManager.AddComment("C8", "原材料成本上升导致毛利率下降", "成本分析师");
        _commentManager.AddComment("C9", "人工成本控制效果显著", "人力资源部");
        
        // 利润相关注释
        _commentManager.AddComment("D12", "净利润增长符合预期", "财务总监");
        _commentManager.AddComment("D13", "投资收益贡献显著", "投资部");
        
        // 特殊项目注释
        var specialFormat = new CommentFormat
        {
            BackgroundColor = RGB(255, 200, 200), // 浅红色背景
            BorderColor = RGB(255, 0, 0), // 红色边框
            FontSize = 11
        };
        
        _commentManager.AddComment("E15", "注意：此项包含一次性重组费用", "审计部");
        _commentManager.FormatComment("E15", specialFormat);
        
        // 预测数据注释
        var forecastFormat = new CommentFormat
        {
            BackgroundColor = RGB(200, 255, 200), // 浅绿色背景
            BorderColor = RGB(0, 128, 0), // 深绿色边框
            FontSize = 10
        };
        
        _commentManager.AddComment("F20", "基于市场趋势的乐观预测", "市场部");
        _commentManager.FormatComment("F20", forecastFormat);
    }
    
    /// <summary>
    /// 添加审计意见
    /// </summary>
    public void AddAuditComment(string cellAddress, string auditText)
    {
        var auditFormat = new CommentFormat
        {
            BackgroundColor = RGB(255, 255, 200), // 浅黄色背景
            BorderColor = RGB(255, 165, 0), // 橙色边框
            FontSize = 12
        };
        
        _commentManager.AddComment(cellAddress, $"审计意见: {auditText}", "审计部");
        _commentManager.FormatComment(cellAddress, auditFormat);
    }
    
    /// <summary>
    /// 导出所有注释用于审计
    /// </summary>
    public void ExportCommentsForAudit(string filePath)
    {
        _commentManager.ExportCommentsToFile(filePath);
    }
    
    /// <summary>
    /// 显示所有重要注释
    /// </summary>
    public void ShowImportantComments()
    {
        var comments = _commentManager.GetAllComments();
        
        foreach (var comment in comments)
        {
            if (comment.Text.Contains("注意") || comment.Text.Contains("审计") || 
                comment.Text.Contains("重要") || comment.Author == "审计部")
            {
                _commentManager.ToggleCommentVisibility(comment.CellAddress, true);
            }
        }
    }
    
    private int RGB(int red, int green, int blue)
    {
        return (red & 0xFF) | ((green & 0xFF) << 8) | ((blue & 0xFF) << 16);
    }
}
```

## 合并单元格操作

### 合并单元格基础操作

合并单元格是创建复杂表格布局的重要技术，但需要谨慎使用以避免数据处理问题。

```csharp
using MudTools.OfficeInterop.Excel;
using System;
using System.Collections.Generic;

/// <summary>
/// 合并单元格管理器
/// 提供安全的合并单元格操作功能
/// </summary>
public class MergeCellManager
{
    private readonly IExcelWorksheet _worksheet;
    
    public MergeCellManager(IExcelWorksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }
    
    /// <summary>
    /// 合并单元格
    /// </summary>
    public void MergeCells(string rangeAddress, string text = "", bool centerText = true)
    {
        var range = _worksheet.Range(rangeAddress);
        if (range == null) return;
        
        // 检查是否已经是合并单元格
        if (range.MergeCells)
        {
            Console.WriteLine($"区域 {rangeAddress} 已经是合并单元格");
            return;
        }
        
        // 设置文本
        if (!string.IsNullOrEmpty(text))
        {
            range.Value = text;
        }
        
        // 合并单元格
        range.Merge();
        
        // 设置文本居中
        if (centerText)
        {
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = XlVAlign.xlVAlignCenter;
        }
        
        // 设置边框
        range.Borders.LineStyle = XlLineStyle.xlContinuous;
        range.Borders.Weight = XlBorderWeight.xlThin;
    }
    
    /// <summary>
    /// 取消合并单元格
    /// </summary>
    public void UnmergeCells(string rangeAddress)
    {
        var range = _worksheet.Range(rangeAddress);
        if (range == null) return;
        
        if (range.MergeCells)
        {
            range.UnMerge();
        }
    }
    
    /// <summary>
    /// 检查是否为合并单元格
    /// </summary>
    public bool IsMerged(string cellAddress)
    {
        var range = _worksheet.Range(cellAddress);
        return range?.MergeCells ?? false;
    }
    
    /// <summary>
    /// 获取合并单元格的范围
    /// </summary>
    public string GetMergedRange(string cellAddress)
    {
        var range = _worksheet.Range(cellAddress);
        if (range == null) return string.Empty;
        
        if (range.MergeCells)
        {
            return range.MergeArea.Address;
        }
        
        return cellAddress; // 单个单元格
    }
    
    /// <summary>
    /// 安全地设置合并单元格的值
    /// </summary>
    public void SetMergedCellValue(string rangeAddress, object value)
    {
        var range = _worksheet.Range(rangeAddress);
        if (range == null) return;
        
        // 如果是合并单元格，只在第一个单元格设置值
        if (range.MergeCells)
        {
            var mergeArea = range.MergeArea;
            if (mergeArea != null)
            {
                mergeArea.Cells[1, 1].Value = value;
            }
        }
        else
        {
            range.Value = value;
        }
    }
    
    /// <summary>
    /// 获取合并单元格的值
    /// </summary>
    public object GetMergedCellValue(string cellAddress)
    {
        var range = _worksheet.Range(cellAddress);
        if (range == null) return null;
        
        // 如果是合并单元格，返回第一个单元格的值
        if (range.MergeCells)
        {
            var mergeArea = range.MergeArea;
            return mergeArea?.Cells[1, 1].Value;
        }
        
        return range.Value;
    }
    
    /// <summary>
    /// 批量合并标题行
    /// </summary>
    public void BatchMergeTitles(List<TitleMergeConfig> configs)
    {
        foreach (var config in configs)
        {
            MergeCells(config.RangeAddress, config.Text, config.CenterText);
            
            // 设置标题格式
            var range = _worksheet.Range(config.RangeAddress);
            if (range != null)
            {
                range.Font.Bold = true;
                range.Font.Size = config.FontSize;
                
                if (config.BackgroundColor.HasValue)
                {
                    range.Interior.Color = config.BackgroundColor.Value;
                }
            }
        }
    }
    
    /// <summary>
    /// 创建复杂表格布局
    /// </summary>
    public void CreateComplexTableLayout()
    {
        // 主标题
        MergeCells("A1:F1", "年度财务报告", true);
        SetTitleFormat("A1:F1", 16, RGB(200, 200, 255));
        
        // 副标题
        MergeCells("A2:F2", $"{DateTime.Now.Year}年度", true);
        SetTitleFormat("A2:F2", 12, RGB(230, 230, 255));
        
        // 收入部分标题
        MergeCells("A4:C4", "收入分析", true);
        SetSectionHeaderFormat("A4:C4");
        
        // 成本部分标题
        MergeCells("D4:F4", "成本分析", true);
        SetSectionHeaderFormat("D4:F4");
        
        // 季度标题
        for (int i = 0; i < 4; i++)
        {
            int row = 5 + i * 3;
            MergeCells($"A{row}:A{row + 2}", $"第{i + 1}季度", true);
        }
    }
    
    private void SetTitleFormat(string rangeAddress, int fontSize, int backgroundColor)
    {
        var range = _worksheet.Range(rangeAddress);
        if (range != null)
        {
            range.Font.Bold = true;
            range.Font.Size = fontSize;
            range.Font.Color = RGB(0, 0, 128); // 深蓝色
            range.Interior.Color = backgroundColor;
        }
    }
    
    private void SetSectionHeaderFormat(string rangeAddress)
    {
        var range = _worksheet.Range(rangeAddress);
        if (range != null)
        {
            range.Font.Bold = true;
            range.Font.Size = 12;
            range.Interior.Color = RGB(240, 240, 240); // 浅灰色
        }
    }
    
    // 辅助方法：RGB颜色转换
    private int RGB(int red, int green, int blue)
    {
        return (red & 0xFF) | ((green & 0xFF) << 8) | ((blue & 0xFF) << 16);
    }
}

/// <summary>
/// 标题合并配置
/// </summary>
public class TitleMergeConfig
{
    public string RangeAddress { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public bool CenterText { get; set; } = true;
    public int FontSize { get; set; } = 12;
    public int? BackgroundColor { get; set; }
}

// 对齐方式枚举（简化版）
public enum XlHAlign
{
    xlHAlignLeft = 1,
    xlHAlignCenter = 2,
    xlHAlignRight = 3
}

public enum XlVAlign
{
    xlVAlignTop = 1,
    xlVAlignCenter = 2,
    xlVAlignBottom = 3
}

public enum XlLineStyle
{
    xlContinuous = 1
}

public enum XlBorderWeight
{
    xlThin = 1
}
```

## 主题和模板应用

### 主题管理基础

主题是确保文档视觉一致性的重要工具，可以统一字体、颜色和效果设置。

```csharp
using MudTools.OfficeInterop.Excel;
using System;
using System.Collections.Generic;

/// <summary>
/// 主题管理器
/// 提供统一的主题应用和管理功能
/// </summary>
public class ThemeManager
{
    private readonly IExcelWorkbook _workbook;
    
    public ThemeManager(IExcelWorkbook workbook)
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
    }
    
    /// <summary>
    /// 应用公司标准主题
    /// </summary>
    public void ApplyCorporateTheme()
    {
        var corporateTheme = new CorporateTheme();
        ApplyTheme(corporateTheme);
    }
    
    /// <summary>
    /// 应用财务报告主题
    /// </summary>
    public void ApplyFinancialTheme()
    {
        var financialTheme = new FinancialTheme();
        ApplyTheme(financialTheme);
    }
    
    /// <summary>
    /// 应用销售报告主题
    /// </summary>
    public void ApplySalesTheme()
    {
        var salesTheme = new SalesTheme();
        ApplyTheme(salesTheme);
    }
    
    /// <summary>
    /// 应用自定义主题
    /// </summary>
    public void ApplyTheme(ExcelTheme theme)
    {
        // 这里可以实现主题应用逻辑
        // 包括字体、颜色、效果等的统一设置
        
        Console.WriteLine($"应用主题: {theme.Name}");
    }
    
    /// <summary>
    /// 创建自定义主题
    /// </summary>
    public ExcelTheme CreateCustomTheme(string name, ThemeColors colors, ThemeFonts fonts)
    {
        return new ExcelTheme
        {
            Name = name,
            Colors = colors,
            Fonts = fonts
        };
    }
}

/// <summary>
/// Excel主题定义
/// </summary>
public class ExcelTheme
{
    public string Name { get; set; } = string.Empty;
    public ThemeColors Colors { get; set; } = new ThemeColors();
    public ThemeFonts Fonts { get; set; } = new ThemeFonts();
}

/// <summary>
/// 主题颜色
/// </summary>
public class ThemeColors
{
    public int PrimaryColor { get; set; } = RGB(0, 112, 192); // 蓝色
    public int SecondaryColor { get; set; } = RGB(237, 125, 49); // 橙色
    public int AccentColor { get; set; } = RGB(112, 173, 71); // 绿色
    public int TextColor { get; set; } = RGB(0, 0, 0); // 黑色
    public int BackgroundColor { get; set; } = RGB(255, 255, 255); // 白色
}

/// <summary>
/// 主题字体
/// </summary>
public class ThemeFonts
{
    public string HeadingFont { get; set; } = "微软雅黑";
    public string BodyFont { get; set; } = "宋体";
    public int HeadingSize { get; set; } = 14;
    public int BodySize { get; set; } = 11;
}

/// <summary>
/// 公司标准主题
/// </summary>
public class CorporateTheme : ExcelTheme
{
    public CorporateTheme()
    {
        Name = "公司标准主题";
        Colors = new ThemeColors
        {
            PrimaryColor = RGB(0, 84, 159), // 公司蓝色
            SecondaryColor = RGB(255, 204, 0), // 公司黄色
            AccentColor = RGB(132, 189, 0) // 公司绿色
        };
        Fonts = new ThemeFonts
        {
            HeadingFont = "微软雅黑",
            BodyFont = "宋体",
            HeadingSize = 16,
            BodySize = 11
        };
    }
}

/// <summary>
/// 财务报告主题
/// </summary>
public class FinancialTheme : ExcelTheme
{
    public FinancialTheme()
    {
        Name = "财务报告主题";
        Colors = new ThemeColors
        {
            PrimaryColor = RGB(0, 112, 192), // 专业蓝色
            SecondaryColor = RGB(165, 165, 165), // 中性灰色
            AccentColor = RGB(217, 83, 79) // 警示红色
        };
        Fonts = new ThemeFonts
        {
            HeadingFont = "Times New Roman",
            BodyFont = "Arial",
            HeadingSize = 14,
            BodySize = 10
        };
    }
}

/// <summary>
/// 销售报告主题
/// </summary>
public class SalesTheme : ExcelTheme
{
    public SalesTheme()
    {
        Name = "销售报告主题";
        Colors = new ThemeColors
        {
            PrimaryColor = RGB(237, 125, 49), // 活力橙色
            SecondaryColor = RGB(112, 173, 71), // 成功绿色
            AccentColor = RGB(91, 155, 213) // 信任蓝色
        };
        Fonts = new ThemeFonts
        {
            HeadingFont = "Arial",
            BodyFont = "Calibri",
            HeadingSize = 15,
            BodySize = 11
        };
    }
}

// 辅助方法：RGB颜色转换
public static int RGB(int red, int green, int blue)
{
    return (red & 0xFF) | ((green & 0xFF) << 8) | ((blue & 0xFF) << 16);
}
```

## 总结

高级格式化技巧是Excel自动化开发中的核心能力，通过合理运用数据验证、超链接、注释管理和合并单元格等技术，可以创建出专业、美观且功能丰富的业务文档。

### 关键技术要点

1. **数据验证**：确保数据质量，防止无效输入
2. **超链接管理**：创建交互式文档，提升用户体验
3. **注释系统**：增强文档可读性和协作性
4. **合并单元格**：创建复杂表格布局，但需谨慎使用
5. **主题应用**：确保视觉一致性，提升专业形象

### 最佳实践建议

- 合理使用数据验证，避免过度限制用户输入
- 超链接应指向稳定可靠的资源
- 注释内容应简洁明了，避免信息过载
- 合并单元格主要用于标题和布局，避免在数据区域使用
- 主题设计应符合企业形象和业务需求

通过掌握这些高级格式化技巧，开发者可以创建出既美观又实用的Excel自动化解决方案，满足各种复杂的业务需求。
# 第11章：文档保护和安全

在处理敏感文档时，保护文档内容和限制编辑权限是非常重要的。MudTools.OfficeInterop.Word库提供了完整的文档保护功能，包括密码保护、编辑限制、内容保护等。本章将详细介绍如何使用这些功能确保文档的安全性和完整性。

## 密码保护

密码保护是最基本的文档安全措施，可以防止未授权用户打开或修改文档。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 添加示例内容
document.Range().Text = "这是受保护的敏感文档内容。\n包含重要的商业信息。";

// 设置打开密码
document.Password = "OpenPassword123";
```

设置打开文档所需的密码。

```csharp
// 设置修改密码（编辑密码）
document.WritePassword = "EditPassword456";
```

设置修改文档内容所需的密码。

```csharp
// 设置密码加密选项
document.EncryptionProvider = "Microsoft Enhanced RSA and AES Cryptographic Provider";
```

设置加密提供程序。

```csharp
// 另存为受密码保护的文档
document.SaveAs2(
    FileName: @"C:\temp\ProtectedDocument.docx",
    Password: "OpenPassword123",
    WritePassword: "EditPassword456"
);

Console.WriteLine("文档已使用密码保护");
```

保存受密码保护的文档。

## 编辑限制

编辑限制可以控制用户对文档特定部分的编辑权限，而不需要完全加密文档。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 添加示例内容
var range = document.Range();
range.Text = "文档标题\n\n这是可以编辑的内容区域。\n\n这是受保护的内容区域，不能编辑。\n\n这是另一个可编辑区域。";

// 定义可编辑区域
var editableRange1 = document.Range(15, 27); // "这是可以编辑的内容区域。"
var editableRange2 = document.Range(65, 78); // "这是另一个可编辑区域。"
```

定义文档中的可编辑区域范围。

```csharp
// 添加可编辑区域
var ed1 = document.EditableRanges.Add(editableRange1);
ed1.Editors.Add(WdEditorType.wdEditorEveryone); // 所有人可编辑

var ed2 = document.EditableRanges.Add(editableRange2);
ed2.Editors.Add("特定用户组"); // 特定用户组可编辑
```

添加可编辑区域并设置编辑权限：
- wdEditorEveryone：所有人都可以编辑
- 指定特定用户组名称

```csharp
// 应用编辑限制
document.Protect(
    Type: WdProtectionType.wdAllowOnlyReading, // 只读保护
    NoReset: true,
    Password: "ProtectionPass123"
);

Console.WriteLine("编辑限制已应用");
```

应用文档保护：
- Type = WdProtectionType.wdAllowOnlyReading：只读保护
- NoReset = true：不重置现有保护
- Password：保护密码

## 内容保护

内容保护可以保护文档的特定部分，如表单字段、书签区域等。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 创建表单文档
var range = document.Range();
range.Text = "员工信息表\n\n姓名：________________\n部门：________________\n职位：________________\n薪资：________________";

// 添加书签保护
range.Collapse(WdCollapseDirection.wdCollapseStart);
range.Text = "【受保护内容开始】";
range.Collapse(WdCollapseDirection.wdCollapseEnd);
```

创建表单文档并添加书签保护标记。

```csharp
// 添加一些内容
range.Text = "\n\n机密信息：这部分内容受到保护。";
var confidentialRange = document.Range(range.Start - 20, range.End);

// 为机密内容添加书签
var bookmark = document.Bookmarks.Add("ConfidentialSection", confidentialRange);
```

为机密内容添加书签。

```csharp
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
formField.TextInput.EditType(WdTextInputType.wdRegularText, "默认值", true); // 只读
```

添加受保护的表单字段并设置为只读。

```csharp
// 应用保护
document.Protect(
    Type: WdProtectionType.wdAllowOnlyFormFields, // 仅允许表单字段编辑
    NoReset: false,
    Password: "FormPass456"
);

Console.WriteLine("内容保护已应用");
```

应用表单字段保护。

## 数字签名

数字签名可以确保文档的完整性和来源可信性。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 添加文档内容
document.Range().Text = "这是一份需要数字签名的重要合同。\n\n签署方：\n甲方：______________\n乙方：______________\n日期：____年____月____日";

// 检查是否有可用的签名提供商
var signatureProviders = app.SignatureProviders;
Console.WriteLine($"可用签名提供商数量: {signatureProviders.Count}");
```

检查可用的数字签名提供商。

```csharp
// 添加数字签名（需要有效的证书）
try
{
    var signatures = document.Signatures;
    
    // 添加签名行
    var signatureLine = document.Shapes.AddSignatureLine(
        new object(), // SignatureLineSpec
        100, 100, 200, 50);
    
    var sigLine = signatureLine.SignatureLine;
    sigLine.SuggestedSigner = "张三";
    sigLine.SuggestedSignerLine2 = "ABC公司总经理";
    sigLine.SuggestedSignerEmail = "zhangsan@abc.com";
```

添加签名行并设置签名者信息。

```csharp
    // 保存文档以准备签名
    document.SaveAs2(@"C:\temp\DocumentToSign.docx");
    
    Console.WriteLine("签名行已添加，文档已保存");
}
catch (Exception ex)
{
    Console.WriteLine($"添加数字签名时出错: {ex.Message}");
    Console.WriteLine("请确保系统中有有效的数字证书");
}
```

## 文档权限管理

可以通过权限管理设置更细粒度的访问控制。

```csharp
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
```

检查并启用权限管理功能。

```csharp
        // 添加用户权限
        var userPermission = document.Permission.Add(
            "user@example.com",
            MsoPermission.msoPermissionRead + MsoPermission.msoPermissionEdit);
```

为用户添加读取和编辑权限。

```csharp
        // 设置权限到期时间
        userPermission.ExpirationDate = DateTime.Now.AddDays(30);
        
        // 设置所有者权限
        document.Permission.ApplyPolicy(@"C:\policies\CorporatePolicy.xml");
        
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
```

## 实际应用示例

以下示例演示了如何创建一个具有多层次安全保护的商业合同文档：

```csharp
using MudTools.OfficeInterop;
using System;

class SecureDocumentSystem
{
    public static void CreateSecureContract()
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
```

设置文档基本属性。

```csharp
            // 创建合同标题
            var titleRange = document.Range();
            titleRange.Text = "保密协议\n";
            titleRange.Font.Name = "微软雅黑";
            titleRange.Font.Size = 18;
            titleRange.Font.Bold = 1;
            titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            titleRange.ParagraphFormat.SpaceAfter = 24;
```

设置合同标题格式。

```csharp
            // 添加合同正文
            var contentRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            contentRange.Text = "本协议由以下双方于____年____月____日签署：\n\n";
            contentRange.Font.Name = "宋体";
            contentRange.Font.Size = 12;
            
            // 甲方信息（可编辑）
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "甲方（披露方）：\n";
            contentRange.Font.Bold = 1;
            
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "公司名称：________________________\n";
            contentRange.Font.Bold = 0;
            
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "地址：___________________________\n";
```

添加合同正文内容。

```csharp
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
            contentRange.Font.Bold = 1;
```

添加乙方信息。

```csharp
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "公司名称：________________________\n";
            contentRange.Font.Bold = 0;
            
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
```

继续添加合同内容。

```csharp
            // 合同条款（受保护）
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "第一条 保密信息的定义\n";
            contentRange.Font.Bold = 1;
            
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "1.1 保密信息指甲方提供给乙方的任何技术、商业或其他信息...\n\n";
            contentRange.Font.Bold = 0;
            
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "第二条 保密义务\n";
            contentRange.Font.Bold = 1;
            
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "2.1 乙方应对保密信息严格保密...\n\n";
            contentRange.Font.Bold = 0;
            
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "第三条 保密期限\n";
            contentRange.Font.Bold = 1;
            
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "3.1 本协议的保密期限为【    】年...\n\n";
            contentRange.Font.Bold = 0;
```

添加合同条款。

```csharp
            // 添加签名区域
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "\n\n【以下无正文】\n\n";
            contentRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            
            // 添加签名行
            contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            contentRange.Text = "\n\n甲方签字：______________    乙方签字：______________\n";
            contentRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
```

添加签名区域。

```csharp
            // 设置编辑限制 - 只允许在特定区域编辑
            // 定义可编辑区域（签名区域）
            var editableRange = document.Range(
                document.Content.End - 50, // 签名区域开始
                document.Content.End        // 文档末尾
            );
            
            // 添加可编辑区域
            var editableRangeObj = document.EditableRanges.Add(editableRange);
            editableRangeObj.Editors.Add(WdEditorType.wdEditorEveryone);
```

设置编辑限制，只允许在签名区域编辑。

```csharp
            // 应用保护
            document.Protect(
                Type: WdProtectionType.wdAllowOnlyReading,
                NoReset: true,
                Password: "Contract2025"
            );
            
            // 设置打开密码和修改密码
            document.Password = "OpenSecure123";
            document.WritePassword = "EditSecure456";
            
            // 保存文档
            document.SaveAs2(@"C:\temp\SecureContract.docx");
            
            Console.WriteLine("安全合同文档已创建: SecureContract.docx");
            Console.WriteLine("文档已应用以下保护措施：");
            Console.WriteLine("1. 打开密码: OpenSecure123");
            Console.WriteLine("2. 修改密码: EditSecure456");
            Console.WriteLine("3. 编辑限制: 仅允许在签名区域编辑");
            Console.WriteLine("4. 保护密码: Contract2025");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"创建安全文档时出错: {ex.Message}");
        }
    }
```

应用多重保护并保存文档。

```csharp
    public static void VerifyDocumentProtection(var document)
    {
        Console.WriteLine("=== 文档保护状态检查 ===");
        
        // 检查文档是否受保护
        try
        {
            bool isProtected = document.ProtectionType != WdProtectionType.wdNoProtection;
            Console.WriteLine($"文档是否受保护: {isProtected}");
            
            if (isProtected)
            {
                Console.WriteLine($"保护类型: {document.ProtectionType}");
            }
```

验证文档保护状态。

```csharp
            // 检查密码设置
            bool hasPassword = !string.IsNullOrEmpty(document.Password);
            bool hasWritePassword = !string.IsNullOrEmpty(document.WritePassword);
            Console.WriteLine($"是否设置打开密码: {hasPassword}");
            Console.WriteLine($"是否设置修改密码: {hasWritePassword}");
            
            // 检查可编辑区域
            Console.WriteLine($"可编辑区域数量: {document.EditableRanges.Count}");
            
            // 检查书签
            Console.WriteLine($"书签数量: {document.Bookmarks.Count}");
            
            // 检查签名
            Console.WriteLine($"数字签名数量: {document.Signatures.Count}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"检查文档保护状态时出错: {ex.Message}");
        }
    }
}
```

检查各种保护措施的状态。

## 应用场景

1. **合同文档**：保护商业合同中的关键条款和签名区域
2. **财务报告**：限制对敏感财务数据的访问和修改
3. **法律文件**：确保法律文档的完整性和不可篡改性
4. **内部政策**：控制公司内部政策文档的访问权限

## 要点总结

- 密码保护可以防止未授权用户打开或修改文档
- 编辑限制允许控制用户对文档特定部分的编辑权限
- 内容保护可以保护文档的特定区域，如表单字段、书签等
- 数字签名确保文档的完整性和来源可信性
- 权限管理提供更细粒度的访问控制
- 可以组合使用多种保护措施以实现多层次安全

掌握文档保护和安全功能对于处理敏感信息的Word文档至关重要，这些功能使开发者能够创建安全可靠的文档处理系统。
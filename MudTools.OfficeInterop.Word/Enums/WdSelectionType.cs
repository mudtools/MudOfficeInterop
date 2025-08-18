//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 选择范围类型枚举
/// 用于标识当前在文档中选择的内容类型
/// </summary>
public enum WdSelectionType
{
    /// <summary>
    /// 无选择 - 没有选择任何内容
    /// </summary>
    wdNoSelection = 0,
    
    /// <summary>
    /// 插入点 - 光标位置，未选择任何内容
    /// </summary>
    wdSelectionIP = 1,
    
    /// <summary>
    /// 普通选择 - 选择了文本内容
    /// </summary>
    wdSelectionNormal = 2,
    
    /// <summary>
    /// 列选择 - 在表格中选择了一列
    /// </summary>
    wdSelectionColumn = 3,
    
    /// <summary>
    /// 行选择 - 在表格中选择了一行
    /// </summary>
    wdSelectionRow = 4,
    
    /// <summary>
    /// 块选择 - 选择了矩形区域的文本块
    /// </summary>
    wdSelectionBlock = 5
}
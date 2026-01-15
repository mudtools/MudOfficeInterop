//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 中的 XPath 对象，用于处理 XML 数据映射。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelXPath : IOfficeObject<IExcelXPath, MsExcel.XPath>, IDisposable
{
    /// <summary>
    /// 获取所属的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取 XPath 表达式的字符串值。
    /// </summary>
    string Value { get; }

    /// <summary>
    /// 获取一个值，该值指示 XPath 是否指向重复元素（多个匹配项）。
    /// </summary>
    bool Repeating { get; }

    /// <summary>
    /// 设置 XPath 的值及相关属性。
    /// </summary>
    /// <param name="map">与 XPath 关联的 XML 映射对象。</param>
    /// <param name="xPath">XPath 表达式字符串。</param>
    /// <param name="selectionNamespace">XPath 表达式的命名空间 URI（可选）。</param>
    /// <param name="repeating">指示 XPath 是否指向重复元素的布尔值（可选）。</param>
    void SetValue(IExcelXmlMap map, string xPath, string? selectionNamespace = null, bool? repeating = null);

    /// <summary>
    /// 清除 XPath 对象的内容。
    /// </summary>
    void Clear();
}
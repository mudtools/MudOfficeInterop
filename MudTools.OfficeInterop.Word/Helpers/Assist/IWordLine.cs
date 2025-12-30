//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 wdTextRectangle 类型的 Rectangle 对象中的单行文本。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordLine : IOfficeObject<IWordLine>, IDisposable
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取代表对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取指示行是文本行还是表格行的 WdLineType 常量。
    /// </summary>
    WdLineType LineType { get; }

    /// <summary>
    /// 获取或设置表示指定行水平位置的整数（以磅为单位）。
    /// </summary>
    int Left { get; }

    /// <summary>
    /// 获取或设置指定行的垂直位置（以磅为单位）。
    /// </summary>
    int Top { get; }

    /// <summary>
    /// 获取或设置指定对象的宽度（以磅为单位）。
    /// </summary>
    int Width { get; }

    /// <summary>
    /// 获取或设置指定对象的高度（以磅为单位）。
    /// </summary>
    int Height { get; }

    /// <summary>
    /// 获取表示指定对象中包含的文档部分的 Range 对象。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取表示文档页面中文本或图形部分的 Rectangles 集合。
    /// </summary>
    IWordRectangles? Rectangles { get; }
}
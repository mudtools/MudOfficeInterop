//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示水平线格式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordHorizontalLineFormat : IOfficeObject<IWordHorizontalLineFormat, MsWord.HorizontalLineFormat>, IDisposable
{

    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取指示创建指定对象的应用程序的 32 位整数。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取表示指定对象的父对象的对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置指定水平线的长度，表示为窗口宽度的百分比。
    /// </summary>
    float PercentWidth { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否绘制不带 3D 阴影的指定水平线。
    /// </summary>
    bool NoShade { get; set; }

    /// <summary>
    /// 获取或设置表示指定水平线对齐方式的 WdHorizontalLineAlignment 常量。
    /// </summary>
    WdHorizontalLineAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置指定 HorizontalLineFormat 对象的宽度类型。
    /// </summary>
    WdHorizontalLineWidthType WidthType { get; set; }
}
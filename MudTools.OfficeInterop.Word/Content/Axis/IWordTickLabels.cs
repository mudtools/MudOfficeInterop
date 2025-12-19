//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 图表刻度标签的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordTickLabels : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取刻度标签名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置刻度标签方向。
    /// </summary>
    XlTickLabelOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置是否自动缩放字体。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoScaleFont { get; set; }

    /// <summary>
    /// 获取或设置刻度标签数字格式。
    /// </summary>
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取或设置是否自动数字格式。
    /// </summary>
    bool NumberFormatLinked { get; set; }

    /// <summary>
    /// 获取或设置数字格式本地化。
    /// </summary>
    object NumberFormatLocal { get; set; }

    /// <summary>
    /// 获取或设置刻度标签读取顺序。
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取刻度标签深度。
    /// </summary>
    int Depth { get; }

    /// <summary>
    /// 获取或设置刻度标签偏移量。
    /// </summary>
    int Offset { get; set; }

    /// <summary>
    /// 获取或设置刻度标签对齐方式。
    /// </summary>
    int Alignment { get; set; }

    /// <summary>
    /// 获取或设置是否启用多级刻度标签。
    /// </summary>
    bool MultiLevel { get; set; }

    /// <summary>
    /// 获取字体格式。
    /// </summary>
    IWordChartFont? Font { get; }

    /// <summary>
    /// 获取格式对象。
    /// </summary>
    IWordChartFormat? Format { get; }

    /// <summary>
    /// 选择刻度标签。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除刻度标签。
    /// </summary>
    void Delete();
}
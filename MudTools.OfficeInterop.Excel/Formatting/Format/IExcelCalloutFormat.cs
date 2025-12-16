//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 标注形状（Callout Shape）格式设置的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.CalloutFormat
/// 用于控制标注的引线类型、角度、边距、起点等。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelCalloutFormat : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Shape）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置标注的类型（如无引线、单引线、角度引线等）。
    /// 使用 <see cref="MsoCalloutType"/> 枚举。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoCalloutType Type { get; set; }

    /// <summary>
    /// 获取或设置标注引线的角度（仅对角度引线有效）。
    /// 使用 <see cref="MsoCalloutAngleType"/> 枚举。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoCalloutAngleType Angle { get; set; }

    /// <summary>
    /// 获取或设置标注形状的边框显示状态。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Border { get; set; }

    /// <summary>
    /// 获取或设置标注文本与连接线之间的间隙大小（单位：磅）。
    /// </summary>
    float Gap { get; set; }

    /// <summary>
    /// 获取或设置标注引线的起点相对于文本框的位置（X 偏移，单位：磅）。
    /// 仅对部分引线类型有效。
    /// </summary>
    float Drop { get; }

    /// <summary>
    /// 获取标注引线起点类型（自动/手动）。
    /// 使用 <see cref="MsoCalloutDropType"/> 枚举。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoCalloutDropType DropType { get; }

    /// <summary>
    /// 获取或设置是否在标注中显示引线（仅对带引线类型有效）。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Accent { get; set; }

    /// <summary>
    /// 获取或设置标注引线是否自动调整以避免遮挡文本。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoAttach { get; set; }

    /// <summary>
    /// 获取标注引线的长度（仅对部分类型有效）。
    /// </summary>
    float Length { get; }

    /// <summary>
    /// 将标注引线长度设置为自动调整模式。
    /// 调用此方法后，引线长度将根据文本框和标注位置自动计算最优长度。
    /// </summary>
    void AutomaticLength();

    /// <summary>
    /// 设置标注引线的自定义垂直偏移量（下垂值）。
    /// </summary>
    /// <param name="Drop">指定引线起点相对于文本框的垂直偏移量，单位为磅。</param>
    void CustomDrop(float Drop);

    /// <summary>
    /// 设置标注引线的自定义长度。
    /// </summary>
    /// <param name="Length">指定引线的长度，单位为磅。</param>
    void CustomLength(float Length);

    /// <summary>
    /// 将标注引线起点设置为预设位置之一。
    /// </summary>
    /// <param name="dropType">指定预设的引线起点位置，使用 <see cref="MsoCalloutDropType"/> 枚举。</param>
    void PresetDrop([ComNamespace("MsCore")] MsoCalloutDropType dropType);
}
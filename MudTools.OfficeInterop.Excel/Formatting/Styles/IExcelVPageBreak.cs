//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel工作表中的垂直分页符接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelVPageBreak : IOfficeObject<IExcelVPageBreak>, IDisposable
{

    /// <summary>
    /// 返回表示 Microsoft Excel 应用程序的 Application 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 返回指定对象的父工作表。
    /// </summary>
    IExcelWorksheet? Parent { get; }

    /// <summary>
    /// 获取或设置分页符的类型。
    /// </summary>
    XlPageBreak Type { get; set; }

    /// <summary>
    /// 获取指定分页符的范围类型：全屏或仅在打印区域内。
    /// </summary>
    XlPageBreakExtent Extent { get; }

    /// <summary>
    /// 获取或设置定义分页符位置的单元格（一个 Range 对象）。
    /// </summary>
    IExcelRange? Location { get; set; }

    /// <summary>
    /// 删除对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 将分页符拖出打印区域。
    /// </summary>
    /// <param name="direction">必需 XlDirection。分页符被拖动的方向。</param>
    /// <param name="regionIndex">必需 Integer。分页符的打印区域索引（如果用户拖动分页符，鼠标指针所在区域的索引）。如果打印区域是连续的，则只有一个打印区域。如果打印区域是不连续的，则有多个打印区域。</param>
    void DragOff(XlDirection direction, int regionIndex);

}
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel工作表中的水平分页符接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelHPageBreak : IOfficeObject<IExcelHPageBreak>, IDisposable
{

    /// <summary>
    /// 获取父级水平分页符集合
    /// </summary>
    IExcelWorksheet? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置水平分页符的类型
    /// </summary>
    XlPageBreak Type { get; set; }

    /// <summary>
    /// 获取或设置水平分页符的位置范围
    /// </summary>
    IExcelRange? Location { get; set; }

    /// <summary>
    /// 获取水平分页符的应用范围类型
    /// </summary>
    XlPageBreakExtent Extent { get; }


    /// <summary>
    /// 将分页符从指定方向和区域索引拖离
    /// </summary>
    /// <param name="direction">拖离的方向</param>
    /// <param name="regionIndex">区域索引</param>
    void DragOff(XlDirection direction, int regionIndex);

    /// <summary>
    /// 移除水平分页符
    /// </summary>
    void Delete();

}
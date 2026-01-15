//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示 Excel 表格（ListObject）中的一行数据，提供对行属性和操作的封装。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelListRow : IOfficeObject<IExcelListRow, MsExcel.ListRow>, IDisposable
{
    /// <summary>
    /// 获取此行所属的父对象（通常是 ListObject）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此行所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取此行在 ListRows 集合中的索引（从 1 开始）。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取一个值，指示此行是否包含无效数据。
    /// </summary>
    /// <value>
    /// 如果行包含无效数据则为 <see langword="true"/>；否则为 <see langword="false"/>。
    /// </value>
    bool InvalidData { get; }

    /// <summary>
    /// 获取此行对应的单元格范围（Range），包含该行所有列的数据。
    /// </summary>
    IExcelRange? Range { get; }

    /// <summary>
    /// 删除此行（将从表格中移除该数据行）。
    /// </summary>
    void Delete();
}
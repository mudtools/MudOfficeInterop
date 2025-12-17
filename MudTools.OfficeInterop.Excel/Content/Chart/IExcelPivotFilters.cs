//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelPivotFilters : IEnumerable<IExcelPivotFilter>, IDisposable
{
    /// <summary>
    /// 获取该对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 <see cref="IExcelApplication"/> 对象，该对象代表 Microsoft Excel 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取集合中对象的数目。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取集合中指定索引位置的 <see cref="IExcelPivotFilter"/> 对象。
    /// </summary>
    /// <param name="index">要获取的对象的索引</param>
    /// <returns>指定索引位置的 <see cref="IExcelPivotFilter"/> 对象</returns>
    IExcelPivotFilter? this[int index] { get; }

    /// <summary>
    /// 获取集合中具有指定名称的 <see cref="IExcelPivotFilter"/> 对象。
    /// </summary>
    /// <param name="name">要获取的对象的名称</param>
    /// <returns>具有指定名称的 <see cref="IExcelPivotFilter"/> 对象</returns>
    IExcelPivotFilter? this[string name] { get; }

    /// <summary>
    /// 向数据透视表中添加一个过滤器。
    /// </summary>
    /// <param name="Type">过滤器类型</param>
    /// <param name="DataField">数据字段</param>
    /// <param name="Value1">第一个值</param>
    /// <param name="Value2">第二个值</param>
    /// <param name="Order">排序顺序</param>
    /// <param name="Name">过滤器名称</param>
    /// <param name="Description">描述信息</param>
    /// <param name="MemberPropertyField">成员属性字段</param>
    /// <param name="WholeDayFilter">是否整日过滤</param>
    /// <returns>添加的数据透视表过滤器</returns>
    IExcelPivotFilter Add2(
        XlPivotFilterType Type,
        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.IDispatch)] object DataField,
        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Struct)] object Value1,
        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Struct)] object Value2,
        [System.Runtime.InteropServices.Optional] object Order,
        [System.Runtime.InteropServices.Optional] object Name,
        [System.Runtime.InteropServices.Optional] object Description,
        [System.Runtime.InteropServices.Optional] object MemberPropertyField,
        [System.Runtime.InteropServices.Optional] object WholeDayFilter);
}

//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 图表分类集合的接口
/// </summary>
/// <remarks>
/// 此接口提供了对 Excel 图表分类集合的访问功能，允许通过索引或名称获取特定的图表分类
/// </remarks>
[ComCollectionWrap(ComNamespace = "MsExcel"), ItemIndex, NoneEnumerable]
public interface IExcelCategoryCollection : IOfficeObject<IExcelCategoryCollection, MsExcel.CategoryCollection>, IEnumerable<IExcelChartCategory?>, IDisposable
{
    /// <summary>
    /// 获取集合中图表分类的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取该对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 通过从1开始的索引获取图表分类
    /// </summary>
    /// <param name="index">要获取的图表分类的从零开始的索引</param>
    /// <returns>指定索引处的图表分类</returns>
    IExcelChartCategory? this[int index] { get; }

    /// <summary>
    /// 通过名称获取图表分类
    /// </summary>
    /// <param name="name">要获取的图表分类的名称</param>
    /// <returns>具有指定名称的图表分类</returns>
    IExcelChartCategory? this[string name] { get; }
}

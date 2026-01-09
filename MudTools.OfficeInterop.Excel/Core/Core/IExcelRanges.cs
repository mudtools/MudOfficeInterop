//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel范围集合的接口，继承自IDisposable和IEnumerable&lt;IExcelRange&gt;接口
/// 用于管理和操作多个Excel范围对象的集合
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelRanges : IDisposable, IOfficeObject<IExcelRanges, MsExcel.Ranges>, IEnumerable<IExcelRange?>
{
    /// <summary>
    /// 获取工作簿集合所在的父对象（通常是Application）
    /// 对应 Workbooks.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取工作簿集合所在的Application对象
    /// 对应 Workbooks.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取范围集合中的范围数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取范围（索引从1开始）
    /// </summary>
    /// <param name="index">范围索引</param>
    /// <returns>范围对象</returns>
    IExcelRange? this[int index] { get; }

    /// <summary>
    /// 根据名称获取范围
    /// </summary>
    /// <param name="name">范围名称</param>
    /// <returns>范围对象</returns>
    IExcelRange? this[string name] { get; }


}
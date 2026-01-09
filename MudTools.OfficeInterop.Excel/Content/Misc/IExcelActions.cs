//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示指定系列的所有 Action 对象的集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelActions : IOfficeObject<IExcelActions, MsExcel.Actions>, IEnumerable<IExcelAction?>, IDisposable
{
    /// <summary>
    /// 获取当前COM对象的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取当前COM对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中的对象数量。只读。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取集合中的 Action 对象。
    /// </summary>
    /// <param name="index">操作的索引值。</param>
    /// <returns>表示工作簿中操作的 Action 对象。</returns>
    IExcelAction? this[int index] { get; }

    /// <summary>
    /// 通过索引获取集合中的 Action 对象。
    /// </summary>
    /// <param name="name">操作的索引值。</param>
    /// <returns>表示工作簿中操作的 Action 对象。</returns>
    IExcelAction? this[string name] { get; }

}
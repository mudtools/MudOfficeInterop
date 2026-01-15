//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 定义Excel页面集合的接口，提供对多个Excel页面的访问和管理功能
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelPages : IOfficeObject<IExcelPages, MsExcel.Pages>, IDisposable, IEnumerable<IExcelPage?>
{
    /// <summary>
    /// 获取页面集合中的页面数量
    /// </summary>
    int? Count { get; }

    /// <summary>
    /// 根据索引获取页面（索引从1开始）
    /// </summary>
    /// <param name="index">页面索引</param>
    /// <returns>页面对象</returns>
    IExcelPage? this[int index] { get; }

    /// <summary>
    /// 根据名称获取页面
    /// </summary>
    /// <param name="name">页面名称</param>
    /// <returns>页面对象</returns>
    IExcelPage? this[string name] { get; }

}
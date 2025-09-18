//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

// =============================================
// 接口定义：列表列集合（ListColumns）
// =============================================
/// <summary>
/// 表示 Excel 表格（ListObject）中所有列的集合，支持遍历和索引访问。
/// </summary>
public interface IExcelListColumns : IEnumerable<IExcelListColumn>, IDisposable
{
    /// <summary>
    /// 获取集合中列的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引（从 1 开始）获取指定的列。
    /// </summary>
    /// <param name="index">列索引（1-based）</param>
    /// <returns>对应的列对象</returns>
    IExcelListColumn this[int index] { get; }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 ListObject）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 向集合中添加一个新列（插入在末尾）。
    /// </summary>
    /// <param name="position">插入位置（可选，默认为末尾）</param>
    /// <returns>新创建的列对象</returns>
    IExcelListColumn Add(int? position = null);
}
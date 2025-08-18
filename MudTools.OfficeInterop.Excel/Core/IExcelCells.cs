//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel工作表中单元格集合的接口，提供对单元格的访问和操作功能
/// 继承自IEnumerable&lt;IExcelRange&gt;和IDisposable接口，支持遍历和资源释放
/// </summary>
public interface IExcelCells : ICoreRange<IExcelRange>, IDisposable
{

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="column">列号</param>
    /// <returns>单元格对象</returns>
    IExcelRange? this[int? row, int? column] { get; }

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="rowAddress">行地址</param>
    /// <param name="columnAddress">列地址</param>
    /// <returns>单元格对象</returns>
    IExcelRange? this[string? rowAddress, string? columnAddress] { get; }

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="address">地址</param>
    IExcelRange? this[string address] { get; }

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <returns>单元格对象</returns>
    IExcelRange? this[int row] { get; }
}
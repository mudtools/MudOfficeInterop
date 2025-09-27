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
public interface IExcelHPageBreak : IDisposable
{
    /// <summary>
    /// 获取或设置水平分页符的类型
    /// </summary>
    XlPageBreak Type { get; set; }

    /// <summary>
    /// 获取或设置水平分页符的位置范围
    /// </summary>
    IExcelRange? Location { get; set; }

    /// <summary>
    /// 获取水平分页符的起始行号
    /// </summary>
    int StartRow { get; }

    /// <summary>
    /// 获取水平分页符的结束行号
    /// </summary>
    int EndRow { get; }

    /// <summary>
    /// 获取水平分页符是否为手动分页符
    /// </summary>
    bool IsManual { get; }

    /// <summary>
    /// 获取水平分页符是否为自动分页符
    /// </summary>
    bool IsAutomatic { get; }

    /// <summary>
    /// 获取父级水平分页符集合
    /// </summary>
    IExcelHPageBreaks? Parent { get; }

    /// <summary>
    /// 获取关联的工作表
    /// </summary>
    IExcelWorksheet? Worksheet { get; }

    /// <summary>
    /// 移除水平分页符
    /// </summary>
    void Delete();

    /// <summary>
    /// 移动分页符到指定行
    /// </summary>
    /// <param name="row">目标行号</param>
    void MoveToRow(int row);

    /// <summary>
    /// 获取分页符前一页的范围
    /// </summary>
    /// <returns>范围对象</returns>
    IExcelRange? GetPreviousPageRange();

    /// <summary>
    /// 获取分页符后一页的范围
    /// </summary>
    /// <returns>范围对象</returns>
    IExcelRange? GetNextPageRange();
}
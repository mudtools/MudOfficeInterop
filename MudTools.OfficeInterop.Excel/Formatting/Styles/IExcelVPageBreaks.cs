//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel工作表中的垂直分页符集合接口
/// </summary>
public interface IExcelVPageBreaks : IDisposable, IEnumerable<IExcelVPageBreak>
{
    /// <summary>
    /// 获取垂直分页符集合中的分页符数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取垂直分页符（索引从1开始）
    /// </summary>
    /// <param name="index">分页符索引</param>
    /// <returns>垂直分页符对象</returns>
    IExcelVPageBreak this[int index] { get; }

    /// <summary>
    /// 添加新的垂直分页符
    /// </summary>
    /// <param name="before">分页符位置（在指定范围之前）</param>
    /// <returns>新创建的垂直分页符对象</returns>
    IExcelVPageBreak Add(IExcelRange before);

    /// <summary>
    /// 根据范围查找垂直分页符
    /// </summary>
    /// <param name="range">范围对象</param>
    /// <returns>垂直分页符对象</returns>
    IExcelVPageBreak FindByRange(IExcelRange range);

    /// <summary>
    /// 根据列号查找垂直分页符
    /// </summary>
    /// <param name="column">列号</param>
    /// <returns>垂直分页符对象</returns>
    IExcelVPageBreak FindByColumn(int column);

    /// <summary>
    /// 移除指定索引的垂直分页符
    /// </summary>
    /// <param name="index">分页符索引</param>
    void RemoveAt(int index);

    /// <summary>
    /// 移除指定范围的垂直分页符
    /// </summary>
    /// <param name="range">范围对象</param>
    void RemoveByRange(IExcelRange range);

    /// <summary>
    /// 移除指定列号的垂直分页符
    /// </summary>
    /// <param name="column">列号</param>
    void RemoveByColumn(int column);

    /// <summary>
    /// 获取父级工作表
    /// </summary>
    IExcelWorksheet Parent { get; }

    /// <summary>
    /// 获取分页符应用的范围
    /// </summary>
    IExcelRange Range { get; }
}
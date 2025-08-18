//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

public interface IExcelHPageBreaks : IDisposable, IEnumerable<IExcelHPageBreak>
{
    /// <summary>
    /// 获取水平分页符集合中的分页符数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取水平分页符（索引从1开始）
    /// </summary>
    /// <param name="index">分页符索引</param>
    /// <returns>水平分页符对象</returns>
    IExcelHPageBreak this[int index] { get; }

    /// <summary>
    /// 添加新的水平分页符
    /// </summary>
    /// <param name="before">分页符位置（在指定范围之前）</param>
    /// <returns>新创建的水平分页符对象</returns>
    IExcelHPageBreak? Add(IExcelRange before);

    /// <summary>
    /// 根据行号查找水平分页符
    /// </summary>
    /// <param name="row">行号</param>
    /// <returns>水平分页符对象</returns>
    IExcelHPageBreak FindByRow(int row);

    /// <summary>
    /// 移除指定索引的水平分页符
    /// </summary>
    /// <param name="index">分页符索引</param>
    void RemoveAt(int index);

    /// <summary>
    /// 移除指定行号的水平分页符
    /// </summary>
    /// <param name="row">行号</param>
    void RemoveByRow(int row);

    /// <summary>
    /// 清除所有水平分页符
    /// </summary>
    void Clear();

    /// <summary>
    /// 获取父级工作表
    /// </summary>
    IExcelWorksheet Parent { get; }

    /// <summary>
    /// 获取分页符应用的范围
    /// </summary>
    IExcelRange Range { get; }

    /// <summary>
    /// 获取指定类型的水平分页符
    /// </summary>
    /// <param name="type">分页符类型</param>
    /// <returns>水平分页符枚举</returns>
    IEnumerable<IExcelHPageBreak> GetPageBreaksByType(XlPageBreak type);

    /// <summary>
    /// 根据位置范围筛选分页符
    /// </summary>
    /// <param name="startRow">起始行</param>
    /// <param name="endRow">结束行</param>
    /// <returns>水平分页符枚举</returns>
    IEnumerable<IExcelHPageBreak> GetPageBreaksInRange(int startRow, int endRow);
}
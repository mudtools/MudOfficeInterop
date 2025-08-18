//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

public interface IExcelVPageBreak : IDisposable
{

    /// <summary>
    /// 获取或设置垂直分页符的类型
    /// </summary>
    XlPageBreak Type { get; set; }

    /// <summary>
    /// 获取或设置垂直分页符的位置范围
    /// </summary>
    IExcelRange Location { get; set; }

    /// <summary>
    /// 获取垂直分页符的起始列号
    /// </summary>
    int StartColumn { get; }

    /// <summary>
    /// 获取垂直分页符的结束列号
    /// </summary>
    int EndColumn { get; }

    /// <summary>
    /// 获取关联的工作表
    /// </summary>
    IExcelWorksheet Worksheet { get; }

    /// <summary>
    /// 获取分页符是否启用
    /// </summary>
    bool Enabled { get; set; }

    /// <summary>
    /// 移除垂直分页符
    /// </summary>
    void Delete();

    /// <summary>
    /// 移动分页符到指定列
    /// </summary>
    /// <param name="column">目标列号</param>
    void MoveToColumn(int column);

    /// <summary>
    /// 获取分页符前一列的范围
    /// </summary>
    /// <returns>范围对象</returns>
    IExcelRange GetPreviousColumnRange();

    /// <summary>
    /// 获取分页符后一列的范围
    /// </summary>
    /// <returns>范围对象</returns>
    IExcelRange GetNextColumnRange();

    /// <summary>
    /// 获取分页符影响的列数
    /// </summary>
    /// <returns>列数</returns>
    int GetAffectedColumnCount();

    /// <summary>
    /// 检查分页符是否与指定范围重叠
    /// </summary>
    /// <param name="range">范围对象</param>
    /// <returns>是否重叠</returns>
    bool OverlapsWith(IExcelRange range);

    /// <summary>
    /// 验证分页符位置是否有效
    /// </summary>
    /// <returns>是否有效</returns>
    bool Validate();

}
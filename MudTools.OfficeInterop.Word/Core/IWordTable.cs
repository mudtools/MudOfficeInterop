//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 表格接口
/// </summary>
public interface IWordTable : IDisposable
{
    /// <summary>
    /// 获取表格行数
    /// </summary>
    int Rows { get; }

    /// <summary>
    /// 获取表格列数
    /// </summary>
    int Columns { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取表格范围
    /// </summary>
    IWordRange Range { get; }

    /// <summary>
    /// 获取指定单元格
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="column">列号</param>
    /// <returns>单元格范围</returns>
    IWordRange Cell(int row, int column);

    /// <summary>
    /// 删除表格
    /// </summary>
    void Delete();

    /// <summary>
    /// 自动调整表格
    /// </summary>
    void AutoFit();

    /// <summary>
    /// 设置表格边框
    /// </summary>
    void SetBorders(bool enable = true);
}

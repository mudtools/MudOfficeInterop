//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel中的表格对象（ListObject）接口，提供对Excel表格的各种操作和属性访问功能。
/// 该接口继承自IDisposable，使用完后需要正确释放资源。
/// </summary>
public interface IExcelListObject : IDisposable
{
    /// <summary>
    /// 获取或设置表格名称
    /// </summary>
    string Name { get; set; }

    IExcelRange Range { get; }

    /// <summary>
    /// 获取表格的数据范围（Range）
    /// </summary>
    IExcelRange DataRange { get; }

    /// <summary>
    /// 获取表格的标题行范围（Range）
    /// </summary>
    IExcelRange HeaderRowRange { get; }

    /// <summary>
    /// 获取表格的总计行范围（Range）
    /// </summary>
    IExcelRange TotalsRowRange { get; }

    /// <summary>
    /// 获取或设置是否显示标题行
    /// </summary>
    bool ShowHeaders { get; set; }

    /// <summary>
    /// 获取或设置是否显示总计行
    /// </summary>
    bool ShowTotals { get; set; }

    /// <summary>
    /// 获取表格所在的 Worksheet 名称
    /// </summary>
    string WorksheetName { get; }

    /// <summary>
    /// 刷新与数据源的链接（如果存在）
    /// </summary>
    void Refresh();

    /// <summary>
    /// 删除表格
    /// </summary>
    void Delete();
}
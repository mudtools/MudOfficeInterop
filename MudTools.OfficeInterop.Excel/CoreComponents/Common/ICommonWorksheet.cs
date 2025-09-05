//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Sheet公共接口。
/// </summary>
public interface ICommonWorksheet : IDisposable
{
    /// <summary>
    /// 获取图表所在的 Excel Application 对象
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置工作表的名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取图表对象的索引位置
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取工作表所在的父对象（通常是工作簿）
    /// </summary>
    object? Parent { get; }

    /// <summary>
    ///  获取工作表所在的父对象名字。
    /// </summary>
    string? ParentName { get; }

    /// <summary>
    /// 获取图表是否被保护
    /// </summary>
    bool IsProtected { get; }

    /// <summary>
    /// 获取或设置工作表是否可见
    /// </summary>
    bool IsVisible { get; set; }

    /// <summary>
    /// 删除工作表
    /// </summary>
    void Delete();

    /// <summary>
    /// 激活对象
    /// </summary>
    void Activate();

    /// <summary>
    /// 选择对象
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 复制对象
    /// </summary>
    void Copy();
}
//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel公共控件。
/// </summary>
public interface IExcelControl
{
    /// <summary>
    /// 获取或设置下拉框的索引
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置下拉框的名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取下拉框的类型
    /// </summary>
    [IgnoreGenerator]
    XlFormControl Type { get; }

    /// <summary>
    /// 获取父级工作表
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置复选框是否启用
    /// </summary>
    bool Enabled { get; set; }



    /// <summary>
    /// 获取或设置列表框的左侧位置
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置列表框的顶部位置
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置列表框的宽度
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置列表框的高度
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 获取或设置列表框是否可见
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置列表框是否锁定
    /// </summary>
    bool Locked { get; set; }
}

//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// Office FileDialog 对象的二次封装接口
/// 提供对 Microsoft.Office.Core.FileDialog 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeFileDialog : IOfficeObject<IOfficeFileDialog, MsCore.FileDialog>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取文件对话框的父对象（通常是 Application）
    /// 对应 FileDialog.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取文件对话框所在的Application对象
    /// 对应 FileDialog.Application 属性
    /// </summary>
    object? Application { get; }

    /// <summary>
    /// 获取文件对话框的类型
    /// 对应 FileDialog.DialogType 属性
    /// </summary>
    MsoFileDialogType DialogType { get; }

    /// <summary>
    /// 获取或设置文件对话框的标题
    /// 对应 FileDialog.Title 属性
    /// </summary>
    string Title { get; set; }

    /// <summary>
    /// 获取或设置文件对话框的初始文件夹路径
    /// 对应 FileDialog.InitialFileName 属性
    /// </summary>
    string InitialFileName { get; set; }

    /// <summary>
    /// 获取或设置文件对话框的初始视图
    /// 对应 FileDialog.InitialView 属性
    /// </summary>
    MsoFileDialogView InitialView { get; set; }

    /// <summary>
    /// 获取或设置文件对话框是否允许多选
    /// 对应 FileDialog.AllowMultiSelect 属性
    /// </summary>
    bool AllowMultiSelect { get; set; }

    /// <summary>
    /// 获取或设置文件对话框的按钮名称（例如，"Import"）
    /// 对应 FileDialog.ButtonName 属性
    /// </summary>
    string ButtonName { get; set; }

    /// <summary>
    /// 获取或设置文件对话框的过滤器索引（选中的过滤器）
    /// 对应 FileDialog.FilterIndex 属性
    /// </summary>
    int FilterIndex { get; set; }

    /// <summary>
    /// 获取文件对话框的选择项
    /// </summary>
    IOfficeSelectedItems? SelectedItems { get; }

    /// <summary>
    /// 获取文件对话框的过滤器集合
    /// 对应 FileDialog.Filters 属性
    /// </summary>
    IOfficeFileDialogFilters? Filters { get; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 显示文件对话框并等待用户操作
    /// 对应 FileDialog.Show 方法
    /// </summary>
    /// <returns>用户操作结果 (例如, Ok = -1, Cancel = 0)</returns>
    int? Show();
    #endregion
}
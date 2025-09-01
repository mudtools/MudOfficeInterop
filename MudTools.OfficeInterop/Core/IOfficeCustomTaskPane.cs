//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// Office CustomTaskPane 对象的二次封装接口
/// 提供对 Microsoft.Office.Core.CustomTaskPane 的安全访问和操作
/// </summary>
public interface IOfficeCustomTaskPane : IDisposable
{
    #region 事件
    /// <summary>
    /// 当自定义任务窗格可见性发生变化时触发
    /// </summary>
    event EventHandler<TaskPaneVisibleStateChangedEventArgs> VisibleStateChanged;

    /// <summary>
    /// 当自定义任务窗格停靠位置发生变化时触发
    /// </summary>
    event EventHandler<TaskPaneDockPositionChangedEventArgs> DockPositionChanged;
    #endregion

    #region 基础属性
    /// <summary>
    /// 获取自定义任务窗格的标题
    /// 对应 CustomTaskPane.Title 属性
    /// </summary>
    string Title { get; }

    /// <summary>
    /// 获取自定义任务窗格所在的Application对象
    /// 对应 CustomTaskPane.Application 属性
    /// </summary>
    object Application { get; } // 使用 object 作为通用占位符

    /// <summary>
    /// 获取自定义任务窗格关联的窗口对象
    /// 对应 CustomTaskPane.Window 属性
    /// </summary>
    object Window { get; } // 使用 object 作为通用占位符，通常是 Excel.Window, Word.Document 等

    /// <summary>
    /// 获取或设置自定义任务窗格是否可见
    /// 对应 CustomTaskPane.Visible 属性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取自定义任务窗格的内容对象（用户控件）
    /// 对应 CustomTaskPane.Content 属性
    /// </summary>
    object ContentControl { get; }

    /// <summary>
    /// 获取或设置自定义任务窗格的停靠位置
    /// 对应 CustomTaskPane.DockPosition 属性
    /// </summary>
    MsoDockPosition DockPosition { get; set; }

    /// <summary>
    /// 获取或设置自定义任务窗格的停靠位置限制
    /// 对应 CustomTaskPane.DockPositionRestrict 属性
    /// </summary>
    MsoDockPositionRestrict DockPositionRestrict { get; set; }
    #endregion

    #region 位置和大小
    /// <summary>
    /// 获取或设置自定义任务窗格的宽度
    /// 对应 CustomTaskPane.Width 属性
    /// </summary>
    int Width { get; set; }

    /// <summary>
    /// 获取或设置自定义任务窗格的高度
    /// 对应 CustomTaskPane.Height 属性
    /// </summary>
    int Height { get; set; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 删除自定义任务窗格
    /// 对应 CustomTaskPane.Delete 方法
    /// </summary>
    void Delete();
    #endregion

    #region 高级功能 (概念性或依赖具体实现)
    /// <summary>
    /// 刷新自定义任务窗格显示
    /// </summary>
    void Refresh();

    /// <summary>
    /// 激活自定义任务窗格中的内容控件（如果支持）
    /// </summary>
    void ActivateContent();
    #endregion
}

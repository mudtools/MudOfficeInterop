//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// PowerPoint Application 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.PowerPoint.Application 的安全访问和操作
/// </summary>
public interface IPowerPointApplication : IOfficeApplication
{
    #region 基础属性
    /// <summary>
    /// 获取是否处于激活状态
    /// 对应 Application.Activated 属性 (如果存在) 或通过其他方式判断
    /// </summary>
    bool IsActive { get; }

    /// <summary>
    /// 获取应用程序是否正在执行命令（忙碌）
    /// 对应 Application.IsSandboxed 属性 (概念不同，占位) 或通过其他方式判断
    /// </summary>
    bool IsBusy { get; }

    #endregion

    #region 核心对象集合和属性
    /// <summary>
    /// 获取演示文稿集合
    /// 对应 Application.Presentations 属性
    /// </summary>
    IPowerPointPresentations Presentations { get; }

    /// <summary>
    /// 获取窗口集合
    /// 对应 Application.Windows 属性
    /// </summary>
    IPowerPointDocumentWindows Windows { get; }

    /// <summary>
    /// 获取当前活动的演示文稿
    /// 对应 Application.ActivePresentation 属性
    /// </summary>
    IPowerPointPresentation ActivePresentation { get; }

    /// <summary>
    /// 获取当前活动的窗口
    /// 对应 Application.ActiveWindow 属性
    /// </summary>
    IPowerPointDocumentWindow ActiveWindow { get; }

    /// <summary>
    /// 获取当前活动的幻灯片
    /// 对应 Application.ActiveWindow.View.Slide (如果 View 是 Normal) 或其他方式
    /// </summary>
    IPowerPointSlide ActiveSlide { get; }

    /// <summary>
    /// 获取当前的选择对象
    /// 对应 Application.ActiveWindow.Selection 属性
    /// </summary>
    IPowerPointSelection Selection { get; }

    /// <summary>
    /// 获取当前视图对象
    /// 对应 Application.ActiveWindow.View 属性
    /// </summary>
    IPowerPointView ActiveView { get; }
    #endregion


    #region 操作方法
    /// <summary>
    /// 执行指定的 PowerPoint 命令 (MSO 命令 ID)
    /// 对应 Application.CommandBars.ExecuteMso 方法
    /// </summary>
    /// <param name="commandId">命令 ID</param>
    void RunCommand(string commandId);


    /// <summary>
    /// 保存所有打开的演示文稿
    /// </summary>
    void SaveAll();
    #endregion

    #region 文件操作
    /// <summary>
    /// 打开一个现有的演示文稿
    /// 对应 Application.Presentations.Open 方法
    /// </summary>
    /// <param name="filename">文件路径</param>
    /// <param name="readOnly">是否只读</param>
    /// <param name="untitled">是否无标题</param>
    /// <param name="withWindow">是否在新窗口中打开</param>
    /// <returns>打开的演示文稿对象</returns>
    IPowerPointPresentation OpenPresentation(string filename, bool readOnly = false, bool untitled = false, bool withWindow = true);

    /// <summary>
    /// 添加一个新的演示文稿
    /// 对应 Application.Presentations.Add 方法
    /// </summary>
    /// <param name="withWindow">是否在新窗口中打开</param>
    /// <returns>新建的演示文稿对象</returns>
    IPowerPointPresentation AddPresentation(bool withWindow = true);
    #endregion

    #region 事件

    /// <summary>
    /// 当演示文稿打开时触发
    /// </summary>
    event PresentationOpenEventHandler PresentationOpen;

    /// <summary>
    /// 当演示文稿关闭前触发
    /// </summary>
    event PresentationBeforeCloseEventHandler PresentationBeforeClose;

    /// <summary>
    /// 当演示文稿保存时触发
    /// </summary>
    event PresentationSaveEventHandler PresentationSave;

    /// <summary>
    /// 当新演示文稿创建时触发
    /// </summary>
    event NewPresentationEventHandler NewPresentation;

    /// <summary>
    /// 当窗口激活时触发
    /// </summary>
    event WindowActivateEventHandler WindowActivate;

    /// <summary>
    /// 当窗口失活时触发
    /// </summary>
    event WindowDeactivateEventHandler WindowDeactivate;

    /// <summary>
    /// 当窗口选择改变时触发
    /// </summary>
    event WindowSelectionChangeEventHandler WindowSelectionChange;


    /// <summary>
    /// 当演示文稿同步时触发
    /// </summary>
    event PresentationSyncEventHandler PresentationSync;

    /// <summary>
    /// 当演示文稿变化时触发
    /// </summary>
    event PresentationChangeEventHandler PresentationChange;

    /// <summary>
    /// 当幻灯片放映开始时触发
    /// </summary>
    event SlideShowBeginEventHandler SlideShowBegin;

    /// <summary>
    /// 当幻灯片放映结束时触发
    /// </summary>
    event SlideShowEndEventHandler SlideShowEnd;

    /// <summary>
    /// 当幻灯片放映下一页时触发
    /// </summary>
    event SlideShowNextSlideEventHandler SlideShowNextSlide;

    #endregion
}
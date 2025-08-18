//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 演示文稿打开事件处理程序
/// </summary>
/// <param name="presentation">演示文稿对象</param>
public delegate void PresentationOpenEventHandler(IPowerPointPresentation presentation);

/// <summary>
/// 演示文稿关闭前事件处理程序
/// </summary>
/// <param name="presentation">演示文稿对象</param>
public delegate void PresentationBeforeCloseEventHandler(IPowerPointPresentation presentation);

/// <summary>
/// 演示文稿保存前事件处理程序
/// </summary>
/// <param name="presentation">演示文稿对象</param>
public delegate void PresentationSaveEventHandler(IPowerPointPresentation presentation);

/// <summary>
/// 新演示文稿事件处理程序
/// </summary>
/// <param name="presentation">演示文稿对象</param>
public delegate void NewPresentationEventHandler(IPowerPointPresentation presentation);

/// <summary>
/// 窗口激活事件处理程序
/// </summary>
/// <param name="presentation">演示文稿对象</param>
/// <param name="wnd">窗口对象</param>
public delegate void WindowActivateEventHandler(IPowerPointPresentation presentation, IPowerPointDocumentWindow wnd);

/// <summary>
/// 窗口失活事件处理程序
/// </summary>
/// <param name="presentation">演示文稿对象</param>
/// <param name="wnd">窗口对象</param>
public delegate void WindowDeactivateEventHandler(IPowerPointPresentation presentation, IPowerPointDocumentWindow wnd);

/// <summary>
/// 窗口选择变化事件处理程序
/// </summary>
/// <param name="sel">选择对象</param>
public delegate void WindowSelectionChangeEventHandler(IPowerPointSelection sel);

/// <summary>
/// 窗口大小变化事件处理程序
/// </summary>
/// <param name="presentation">演示文稿对象</param>
/// <param name="wnd">窗口对象</param>
public delegate void WindowSizeEventHandler(IPowerPointPresentation presentation, IPowerPointDocumentWindow wnd);

/// <summary>
/// 演示文稿同步事件处理程序
/// </summary>
/// <param name="presentation">演示文稿对象</param>
/// <param name="syncEventType">同步事件类型</param>
public delegate void PresentationSyncEventHandler(IPowerPointPresentation presentation, MsoSyncEventType syncEventType);

/// <summary>
/// 演示文稿变化事件处理程序
/// </summary>
public delegate void PresentationChangeEventHandler();

/// <summary>
/// 幻灯片显示开始事件处理程序
/// </summary>
/// <param name="wnd">幻灯片放映窗口</param>
public delegate void SlideShowBeginEventHandler(IPowerPointSlideShowWindow wnd);

/// <summary>
/// 幻灯片显示结束事件处理程序
/// </summary>
/// <param name="wnd">幻灯片放映窗口</param>
public delegate void SlideShowEndEventHandler(IPowerPointPresentation wnd);

/// <summary>
/// 幻灯片放映下一页事件处理程序
/// </summary>
/// <param name="wnd">幻灯片放映窗口</param>
public delegate void SlideShowNextSlideEventHandler(IPowerPointSlideShowWindow wnd);


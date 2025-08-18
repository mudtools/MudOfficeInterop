//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 文档新建事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
public delegate void DocumentNewEventHandler(IWordDocument doc);

/// <summary>
/// 文档内容变化事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
public delegate void DocumentContentChangeEventHandler(IWordDocument doc);

/// <summary>
/// 文档另存为事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
/// <param name="fileName">文件名</param>
/// <param name="cancel">是否取消操作</param>
public delegate void DocumentBeforeSaveAsEventHandler(IWordDocument doc, ref string fileName, ref bool cancel);


/// <summary>
/// 文档打开事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
public delegate void DocumentOpenEventHandler(IWordDocument doc);

/// <summary>
/// 文档关闭前事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
/// <param name="cancel">是否取消关闭</param>
public delegate void DocumentBeforeCloseEventHandler(IWordDocument doc, ref bool cancel);

/// <summary>
/// 文档保存前事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
/// <param name="saveAsUI">是否显示另存为对话框</param>
/// <param name="cancel">是否取消保存</param>
public delegate void DocumentBeforeSaveEventHandler(IWordDocument doc, ref bool saveAsUI, ref bool cancel);

/// <summary>
/// 新文档事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
public delegate void NewDocumentEventHandler(IWordDocument doc);

/// <summary>
/// 窗口激活事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
/// <param name="wnd">窗口对象</param>
public delegate void WindowActivateEventHandler(IWordDocument doc, IWordWindow wnd);

/// <summary>
/// 窗口失活事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
/// <param name="wnd">窗口对象</param>
public delegate void WindowDeactivateEventHandler(IWordDocument doc, IWordWindow wnd);

/// <summary>
/// 文档同步事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
/// <param name="syncEventType">同步事件类型</param>
public delegate void DocumentSyncEventHandler(IWordDocument doc, MsoSyncEventType syncEventType);

/// <summary>
/// 文档变化事件处理程序
/// </summary>
public delegate void DocumentChangeEventHandler();

/// <summary>
/// 邮件合并数据源加载事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
public delegate void MailMergeDataSourceLoadEventHandler(IWordDocument doc);

/// <summary>
/// 邮件合并数据源验证事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
/// <param name="handled">是否已处理</param>
public delegate void MailMergeDataSourceValidateEventHandler(IWordDocument doc, ref bool handled);

/// <summary>
/// 邮件合并向导新数据源事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
public delegate void MailMergeWizardNewDataSourceEventHandler(IWordDocument doc);

/// <summary>
/// 邮件合并向导状态变化事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
/// <param name="wizardState">向导状态</param>
public delegate void MailMergeWizardStateChangeEventHandler(IWordDocument doc, int wizardState);

/// <summary>
/// 窗口选择变化事件处理程序
/// </summary>
/// <param name="sel">选择对象</param>
public delegate void WindowSelectionChangeEventHandler(IWordSelection sel);

/// <summary>
/// 窗口大小变化事件处理程序
/// </summary>
/// <param name="doc">文档对象</param>
/// <param name="wnd">窗口对象</param>
public delegate void WindowSizeEventHandler(IWordDocument doc, IWordWindow wnd);


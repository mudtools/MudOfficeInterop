//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
public partial interface IWordApplication
{
    #region 事件

    /// <summary>
    /// 当文档打开时触发
    /// </summary>
    event DocumentOpenEventHandler DocumentOpen;

    /// <summary>
    /// 当文档关闭前触发
    /// </summary>
    event DocumentBeforeCloseEventHandler DocumentBeforeClose;

    /// <summary>
    /// 当文档保存前触发
    /// </summary>
    event DocumentBeforeSaveEventHandler DocumentBeforeSave;

    /// <summary>
    /// 当窗口激活时触发
    /// </summary>
    event WindowActivateEventHandler WindowActivate;

    /// <summary>
    /// 当窗口失活时触发
    /// </summary>
    event WindowDeactivateEventHandler WindowDeactivate;

    /// <summary>
    /// 当文档同步时触发
    /// </summary>
    event DocumentSyncEventHandler DocumentSync;

    /// <summary>
    /// 当文档变化时触发
    /// </summary>
    event DocumentChangeEventHandler DocumentChange;

    /// <summary>
    /// 当邮件合并数据源打开时触发
    /// </summary>
    event MailMergeDataSourceLoadEventHandler MailMergeDataSourceLoad;

    /// <summary>
    /// 当邮件合并数据源验证时触发
    /// </summary>
    event MailMergeDataSourceValidateEventHandler MailMergeDataSourceValidate;


    /// <summary>
    /// 当窗口选择改变时触发
    /// </summary>
    event WindowSelectionChangeEventHandler WindowSelectionChange;

    /// <summary>
    /// 当窗口大小改变时触发
    /// </summary>
    event WindowSizeEventHandler WindowSize;

    #endregion
}

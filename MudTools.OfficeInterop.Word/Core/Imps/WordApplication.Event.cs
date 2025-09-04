//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
internal partial class WordApplication
{
    private MsWord.ApplicationEvents4_Event _applicationEvent;

    #region 事件字段
    private DocumentOpenEventHandler _documentOpen;
    private DocumentBeforeCloseEventHandler _documentBeforeClose;
    private DocumentBeforeSaveEventHandler _documentBeforeSave;
    private DocumentNewEventHandler _newDocument;
    private WindowActivateEventHandler _windowActivate;
    private WindowDeactivateEventHandler _windowDeactivate;
    private DocumentSyncEventHandler _documentSync;
    private DocumentChangeEventHandler _documentChange;
    private MailMergeDataSourceLoadEventHandler _mailMergeDataSourceLoad;
    private MailMergeDataSourceValidateEventHandler _mailMergeDataSourceValidate;
    private WindowSelectionChangeEventHandler _windowSelectionChange;
    private WindowSizeEventHandler _windowSize;
    #endregion

    private void ConnectEvents()
    {
        // 连接事件处理程序
        _applicationEvent.DocumentOpen += OnDocumentOpen;
        _applicationEvent.NewDocument += OnNewDocument;
        _applicationEvent.DocumentBeforeClose += OnDocumentBeforeClose;
        _applicationEvent.DocumentBeforeSave += OnDocumentBeforeSave;
        _applicationEvent.WindowActivate += OnWindowActivate;
        _applicationEvent.WindowDeactivate += OnWindowDeactivate;
        _applicationEvent.DocumentSync += OnDocumentSync;
        _applicationEvent.DocumentChange += OnDocumentChange;
        _applicationEvent.MailMergeDataSourceLoad += OnMailMergeDataSourceLoad;
        _applicationEvent.MailMergeDataSourceValidate += OnMailMergeDataSourceValidate;
        _applicationEvent.WindowSelectionChange += OnWindowSelectionChange;
        _applicationEvent.WindowSize += OnWindowSize;
    }

    private void DisconnectEvents()
    {
        if (_applicationEvent == null)
            return;

        // 断开事件处理程序
        _applicationEvent.DocumentOpen -= OnDocumentOpen;
        _applicationEvent.NewDocument -= OnNewDocument;
        _applicationEvent.DocumentBeforeClose -= OnDocumentBeforeClose;
        _applicationEvent.DocumentBeforeSave -= OnDocumentBeforeSave;
        _applicationEvent.WindowActivate -= OnWindowActivate;
        _applicationEvent.WindowDeactivate -= OnWindowDeactivate;
        _applicationEvent.DocumentSync -= OnDocumentSync;
        _applicationEvent.DocumentChange -= OnDocumentChange;
        _applicationEvent.MailMergeDataSourceLoad -= OnMailMergeDataSourceLoad;
        _applicationEvent.MailMergeDataSourceValidate -= OnMailMergeDataSourceValidate;
        _applicationEvent.WindowSelectionChange -= OnWindowSelectionChange;
        _applicationEvent.WindowSize -= OnWindowSize;

        _applicationEvent = null;
    }

    #region 事件处理方法
    private void OnDocumentOpen(MsWord.Document doc)
    {
        _documentOpen?.Invoke(new WordDocument(doc));
    }

    private void OnDocumentBeforeClose(MsWord.Document doc, ref bool cancel)
    {
        _documentBeforeClose?.Invoke(new WordDocument(doc), ref cancel);
    }

    private void OnDocumentBeforeSave(MsWord.Document doc, ref bool saveAsUI, ref bool cancel)
    {
        _documentBeforeSave?.Invoke(new WordDocument(doc), ref saveAsUI, ref cancel);
    }

    private void OnNewDocument(MsWord.Document doc)
    {
        _newDocument?.Invoke(new WordDocument(doc));
    }

    private void OnWindowActivate(MsWord.Document doc, MsWord.Window wnd)
    {
        _windowActivate?.Invoke(new WordDocument(doc), new WordWindow(wnd));
    }

    private void OnWindowDeactivate(MsWord.Document doc, MsWord.Window wnd)
    {
        _windowDeactivate?.Invoke(new WordDocument(doc), new WordWindow(wnd));
    }

    private void OnDocumentSync(MsWord.Document doc, MsCore.MsoSyncEventType syncEventType)
    {
        _documentSync?.Invoke(new WordDocument(doc), (MsoSyncEventType)syncEventType);
    }

    private void OnDocumentChange()
    {
        _documentChange?.Invoke();
    }

    private void OnMailMergeDataSourceLoad(MsWord.Document doc)
    {
        _mailMergeDataSourceLoad?.Invoke(new WordDocument(doc));
    }

    private void OnMailMergeDataSourceValidate(MsWord.Document doc, ref bool handled)
    {
        _mailMergeDataSourceValidate?.Invoke(new WordDocument(doc), ref handled);
    }

    private void OnWindowSelectionChange(MsWord.Selection sel)
    {
        _windowSelectionChange?.Invoke(new WordSelection(sel, _activeDocument));
    }

    private void OnWindowSize(MsWord.Document doc, MsWord.Window wnd)
    {
        _windowSize?.Invoke(new WordDocument(doc), new WordWindow(wnd));
    }
    #endregion

    #region 事件实现
    public event DocumentNewEventHandler DocumentNew
    {
        add { _newDocument += value; }
        remove { _newDocument -= value; }
    }

    /// <summary>
    /// 当文档打开时触发
    /// </summary>
    public event DocumentOpenEventHandler DocumentOpen
    {
        add { _documentOpen += value; }
        remove { _documentOpen -= value; }
    }

    /// <summary>
    /// 当文档关闭前触发
    /// </summary>
    public event DocumentBeforeCloseEventHandler DocumentBeforeClose
    {
        add { _documentBeforeClose += value; }
        remove { _documentBeforeClose -= value; }
    }

    /// <summary>
    /// 当文档保存前触发
    /// </summary>
    public event DocumentBeforeSaveEventHandler DocumentBeforeSave
    {
        add { _documentBeforeSave += value; }
        remove { _documentBeforeSave -= value; }
    }

    /// <summary>
    /// 当窗口激活时触发
    /// </summary>
    public event WindowActivateEventHandler WindowActivate
    {
        add { _windowActivate += value; }
        remove { _windowActivate -= value; }
    }

    /// <summary>
    /// 当窗口失活时触发
    /// </summary>
    public event WindowDeactivateEventHandler WindowDeactivate
    {
        add { _windowDeactivate += value; }
        remove { _windowDeactivate -= value; }
    }

    /// <summary>
    /// 当文档同步时触发
    /// </summary>
    public event DocumentSyncEventHandler DocumentSync
    {
        add { _documentSync += value; }
        remove { _documentSync -= value; }
    }

    /// <summary>
    /// 当文档变化时触发
    /// </summary>
    public event DocumentChangeEventHandler DocumentChange
    {
        add { _documentChange += value; }
        remove { _documentChange -= value; }
    }

    /// <summary>
    /// 当邮件合并数据源打开时触发
    /// </summary>
    public event MailMergeDataSourceLoadEventHandler MailMergeDataSourceLoad
    {
        add { _mailMergeDataSourceLoad += value; }
        remove { _mailMergeDataSourceLoad -= value; }
    }

    /// <summary>
    /// 当邮件合并数据源验证时触发
    /// </summary>
    public event MailMergeDataSourceValidateEventHandler MailMergeDataSourceValidate
    {
        add { _mailMergeDataSourceValidate += value; }
        remove { _mailMergeDataSourceValidate -= value; }
    }

    /// <summary>
    /// 当窗口选择改变时触发
    /// </summary>
    public event WindowSelectionChangeEventHandler WindowSelectionChange
    {
        add { _windowSelectionChange += value; }
        remove { _windowSelectionChange -= value; }
    }

    /// <summary>
    /// 当窗口大小改变时触发
    /// </summary>
    public event WindowSizeEventHandler WindowSize
    {
        add { _windowSize += value; }
        remove { _windowSize -= value; }
    }
    #endregion
}

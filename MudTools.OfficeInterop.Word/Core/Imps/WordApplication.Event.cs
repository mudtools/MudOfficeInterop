//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Imps;
namespace MudTools.OfficeInterop.Word.Imps;

internal partial class WordApplication
{
    private MsWord.ApplicationEvents4_Event? _applicationEvent;

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

    internal WordApplication()
    {
        _disposedValue = true;
    }

    /// <summary>
    ///  使用 <see cref="MsWord.Application"/> COM对象初始化当前实例
    /// </summary>
    internal WordApplication(MsWord.Application comObject)
    {
        _application = comObject ?? throw new ArgumentNullException(nameof(comObject));
        _applicationEvent = _application;
        ConnectEvents();
        _disposedValue = false;
    }

    /// <summary>
    /// 创建一个空白文档
    /// </summary>
    /// <returns>新建的文档对象</returns>
    public IWordDocument? BlankDocument()
    {
        try
        {
            if (_application == null)
                throw new ObjectDisposedException(nameof(_application));
            var doc = _application.Documents.Add();
            var wordDoc = new WordDocument(doc);
            return wordDoc;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create blank document.", ex);
        }
    }

    public void Activate()
    {
        if (_application == null)
            throw new ObjectDisposedException(nameof(_application));
        try
        {
            _application?.Activate();
        }
        catch (COMException cx)
        {
            throw new ExcelOperationException("执行Activate操作失败: " + cx.Message, cx);
        }
        catch (Exception ex)
        {
            throw new ExcelOperationException("执行Activate操作失败", ex);
        }
    }

    public IWordDocument CreateFrom(string templatePath)
    {
        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found.", templatePath);

        try
        {
            if (_application == null)
                throw new ObjectDisposedException(nameof(_application));
            var doc = _application.Documents.Add(templatePath);
            var wordDoc = new WordDocument(doc);
            return wordDoc;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create document from template.", ex);
        }
    }


    public IWordDocument? Open(string filePath, bool readOnly = false, string? password = null)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Document file not found.", filePath);

        try
        {
            if (_application == null)
                throw new ObjectDisposedException(nameof(_application));
            if (Documents == null)
                throw new InvalidOperationException("Documents property is null.");

            var wordDoc = Documents.Open(filePath, readOnly: readOnly, passwordDocument: password);
            return wordDoc;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to open document '{filePath}'.", ex);
        }
    }


    public int WindowStateValue
    {
        get
        {
            if (_application == null)
                throw new ObjectDisposedException(nameof(_application));
            return _application.WindowState.ConvertToInt();
        }
        set
        {
            if (_application == null)
                throw new ObjectDisposedException(nameof(_application));
            _application.WindowState = value.EnumConvert<MsWord.WdWindowState>();
        }
    }

    /// <summary>
    /// 创建文件对话框
    /// </summary>
    /// <param name="fileDialogType">文件对话框类型</param>
    /// <returns>文件对话框对象</returns>
    public IOfficeFileDialog? CreateFileDialog(MsoFileDialogType fileDialogType)
    {
        var dialog = _application?.FileDialog[(MsCore.MsoFileDialogType)(int)fileDialogType];
        return dialog != null ? new OfficeFileDialog(dialog) : null;
    }

    private void ConnectEvents()
    {
        if (_applicationEvent == null)
            return;

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
        _windowSelectionChange?.Invoke(new WordSelection(sel));
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

    #region dispose
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            DisconnectEvents();
            if (_application != null)
            {
                Marshal.ReleaseComObject(_application);
                _application = null;
            }
            _disposableList.Dispose();
            _wordDocuments_Documents?.Dispose();
            _wordDocuments_Documents = null;
            _wordWindows_Windows?.Dispose();
            _wordWindows_Windows = null;
            _wordDocument_ActiveDocument?.Dispose();
            _wordDocument_ActiveDocument = null;
            _wordWindow_ActiveWindow?.Dispose();
            _wordWindow_ActiveWindow = null;
            _wordSelection_Selection?.Dispose();
            _wordSelection_Selection = null;
            _wordRecentFiles_RecentFiles?.Dispose();
            _wordRecentFiles_RecentFiles = null;
            _wordTemplate_NormalTemplate?.Dispose();
            _wordTemplate_NormalTemplate = null;
            _wordSystem_System?.Dispose();
            _wordSystem_System = null;
            _wordAutoCorrect_AutoCorrect?.Dispose();
            _wordAutoCorrect_AutoCorrect = null;
            _wordFontNames_FontNames?.Dispose();
            _wordFontNames_FontNames = null;
            _wordFontNames_LandscapeFontNames?.Dispose();
            _wordFontNames_LandscapeFontNames = null;
            _wordFontNames_PortraitFontNames?.Dispose();
            _wordFontNames_PortraitFontNames = null;
            _wordLanguages_Languages?.Dispose();
            _wordLanguages_Languages = null;
            _wordBrowser_Browser?.Dispose();
            _wordBrowser_Browser = null;
            _wordFileConverters_FileConverters?.Dispose();
            _wordFileConverters_FileConverters = null;
            _wordMailingLabel_MailingLabel?.Dispose();
            _wordMailingLabel_MailingLabel = null;
            _wordDialogs_Dialogs?.Dispose();
            _wordDialogs_Dialogs = null;
            _wordCaptionLabels_CaptionLabels?.Dispose();
            _wordCaptionLabels_CaptionLabels = null;
            _wordAutoCaptions_AutoCaptions?.Dispose();
            _wordAutoCaptions_AutoCaptions = null;
            _wordAddIns_AddIns?.Dispose();
            _wordAddIns_AddIns = null;
            _wordTasks_Tasks?.Dispose();
            _wordTasks_Tasks = null;
            _vbeApplication_VBE?.Dispose();
            _vbeApplication_VBE = null;
            _wordListGalleries_ListGalleries?.Dispose();
            _wordListGalleries_ListGalleries = null;
            _wordTemplates_Templates?.Dispose();
            _wordTemplates_Templates = null;
            _wordKeyBindings_KeyBindings?.Dispose();
            _wordKeyBindings_KeyBindings = null;
            _wordOptions_Options?.Dispose();
            _wordOptions_Options = null;
            _wordDictionaries_CustomDictionaries?.Dispose();
            _wordDictionaries_CustomDictionaries = null;
            _officeFileSearch_FileSearch?.Dispose();
            _officeFileSearch_FileSearch = null;
            _wordHangulHanjaConversionDictionaries_HangulHanjaDictionaries?.Dispose();
            _wordHangulHanjaConversionDictionaries_HangulHanjaDictionaries = null;
            _wordMailMessage_MailMessage?.Dispose();
            _wordMailMessage_MailMessage = null;
            _wordEmailOptions_EmailOptions?.Dispose();
            _wordEmailOptions_EmailOptions = null;
            _officeCOMAddIns_COMAddIns?.Dispose();
            _officeCOMAddIns_COMAddIns = null;
            _wordBibliography_Bibliography?.Dispose();
            _wordBibliography_Bibliography = null;
            _wordOMathAutoCorrect_OMathAutoCorrect?.Dispose();
            _wordOMathAutoCorrect_OMathAutoCorrect = null;
            _officeAssistance_Assistance?.Dispose();
            _officeAssistance_Assistance = null;
            _officeSmartArtLayouts_SmartArtLayouts?.Dispose();
            _officeSmartArtLayouts_SmartArtLayouts = null;
            _officeSmartArtQuickStyles_SmartArtQuickStyles?.Dispose();
            _officeSmartArtQuickStyles_SmartArtQuickStyles = null;
            _officeSmartArtColors_SmartArtColors?.Dispose();
            _officeSmartArtColors_SmartArtColors = null;
            _wordUndoRecord_UndoRecord?.Dispose();
            _wordUndoRecord_UndoRecord = null;
            _officePickerDialog_PickerDialog?.Dispose();
            _officePickerDialog_PickerDialog = null;
            _wordProtectedViewWindows_ProtectedViewWindows?.Dispose();
            _wordProtectedViewWindows_ProtectedViewWindows = null;
            _wordProtectedViewWindow_ActiveProtectedViewWindow?.Dispose();
            _wordProtectedViewWindow_ActiveProtectedViewWindow = null;
            _officeLanguageSettings_LanguageSettings?.Dispose();
            _officeLanguageSettings_LanguageSettings = null;
            _officeCommandBars_CommandBars?.Dispose();
            _officeCommandBars_CommandBars = null;
        }

        _disposedValue = true;
    }
    #endregion
}

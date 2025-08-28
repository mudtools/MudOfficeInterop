//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Imp;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 应用程序实现类
/// </summary>
internal class WordApplication : IWordApplication
{
    private MsWord.Application _application;
    private MsWord.ApplicationEvents4_Event _applicationEvent;
    private IWordDocument _activeDocument;
    private bool _disposedValue;
    private IWordWindows _windows;
    private IWordDocuments _documents;
    private IWordSelection _selection;

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

    public WordAppVisibility Visibility
    {
        get => _application.Visible ? WordAppVisibility.Visible : WordAppVisibility.Hidden;
        set => _application.Visible = value == WordAppVisibility.Visible;
    }

    public int Count => _application.Documents.Count;

    public IWordDocument ActiveDocument => _activeDocument;

    public object Parent => _application.Parent;

    public string Version => _application.Version;

    /// <summary>
    /// 获取应用程序的名称
    /// </summary>
    public string Name => _application?.Name;

    public MsoLanguageID Language => (MsoLanguageID)_application.Language;

    public IWordLanguages Languages => new WordLanguages(_application.Languages);

    /// <summary>
    /// 获取应用程序的当前路径
    /// </summary>
    public string Path => _application?.Path;

    public string? Build => _application?.Build;

    /// <summary>
    /// 获取或设置应用程序是否可见
    /// </summary>
    public bool Visible
    {
        get => _application != null && _application.Visible;
        set
        {
            if (_application != null)
                _application.Visible = value;
        }
    }

    public IOfficeLanguageSettings LanguageSettings => new OfficeLanguageSettings(_application.LanguageSettings);

    public IOfficeCommandBars CommandBars => new OfficeCommandBars(_application.CommandBars);

    public WdAlertLevel DisplayAlerts
    {
        get => (WdAlertLevel)_application.DisplayAlerts;
        set => _application.DisplayAlerts = (MsWord.WdAlertLevel)value;
    }

    /// <summary>
    /// 获取或设置窗口状态
    /// </summary>
    public int WindowState
    {
        get => _application != null ? Convert.ToInt32(_application.WindowState) : 0;
        set
        {
            if (_application != null)
                _application.WindowState = (MsWord.WdWindowState)value;
        }
    }

    /// <summary>
    /// 获取或设置应用程序的高度
    /// </summary>
    public float Height
    {
        get => Convert.ToSingle(_application?.Height);
        set
        {
            if (_application != null)
                _application.Height = Convert.ToInt32(value);
        }
    }

    /// <summary>
    /// 获取或设置应用程序的宽度
    /// </summary>
    public float Width
    {
        get => Convert.ToSingle(_application?.Width);
        set
        {
            if (_application != null)
                _application.Width = Convert.ToInt32(value);
        }
    }

    /// <summary>
    /// 获取或设置应用程序的左边距
    /// </summary>
    public float Left
    {
        get => Convert.ToSingle(_application?.Left);
        set
        {
            if (_application != null)
                _application.Left = Convert.ToInt32(value);
        }
    }

    /// <summary>
    /// 获取或设置应用程序的顶边距
    /// </summary>
    public float Top
    {
        get => Convert.ToSingle(_application?.Top);
        set
        {
            if (_application != null)
                _application.Top = Convert.ToInt32(value);
        }
    }

    public bool ScreenUpdating
    {
        get => _application.ScreenUpdating;
        set => _application.ScreenUpdating = value;
    }


    public string UserName
    {
        get => _application.UserName;
        set => _application.UserName = value;
    }

    public IWordSelection Selection
    {
        get
        {
            _selection ??= new WordSelection(_application.Selection, _activeDocument);
            return _selection;
        }
    }

    public IWordDocuments Documents
    {
        get
        {
            _documents ??= new WordDocuments(_application.Documents, this);
            return _documents;
        }
    }

    public IWordWindows Windows
    {
        get
        {
            _windows ??= new WordWindows(_application.Windows, this);
            return _windows;
        }
    }

    public int WindowCount => _application.Windows.Count;

    public IWordWindow? ActiveWindow
    {
        get
        {
            try
            {
                var activeWindow = _application.ActiveWindow;
                return activeWindow != null ? new WordWindow(activeWindow) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    private IWordTemplate? _wordTemplate;

    public IWordTemplate? NormalTemplate
    {
        get
        {
            if (_wordTemplate != null) return _wordTemplate;
            _wordTemplate = new WordTemplate(_application.NormalTemplate);
            return _wordTemplate;
        }
    }

    public IWordDocument this[int index]
    {
        get
        {
            var doc = _application.Documents[index];
            return new WordDocument(doc, this);
        }
    }

    public IWordDocument Select(int index)
    {
        var doc = this[index];
        doc.Activate();
        return doc;
    }

    public IWordWindow GetWindow(int index)
    {
        return Windows.Item(index);
    }

    public IWordWindow NewWindow()
    {
        return Windows.NewWindow();
    }

    public IWordDocument CreateFrom(string templatePath)
    {
        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found.", templatePath);

        try
        {
            var doc = _application.Documents.Add(templatePath);
            var wordDoc = new WordDocument(doc, this);
            MemorizeActiveDocument(wordDoc);
            return wordDoc;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create document from template.", ex);
        }
    }

    public IWordDocument Open(string filePath, bool readOnly = false, string password = null)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Document file not found.", filePath);

        try
        {
            // 正确的参数类型和 ref 修饰符
            object fileNameObj = filePath;
            object confirmConversionsObj = missing;
            object readOnlyObj = readOnly;
            object addToRecentFilesObj = missing;
            object passwordDocumentObj = string.IsNullOrEmpty(password) ? missing : (object)password;
            object passwordTemplateObj = missing;
            object revertObj = missing;
            object writePasswordDocumentObj = missing;
            object writePasswordTemplateObj = missing;
            object formatObj = missing;
            object encodingObj = missing;
            object visibleObj = missing;
            object openAndRepairObj = missing;
            object documentDirectionObj = missing;
            object noEncodingDialogObj = missing;
            object xMLTransformObj = missing;

            var doc = _application.Documents.Open(
                ref fileNameObj,
                ref confirmConversionsObj,
                ref readOnlyObj,
                ref addToRecentFilesObj,
                ref passwordDocumentObj,
                ref passwordTemplateObj,
                ref revertObj,
                ref writePasswordDocumentObj,
                ref writePasswordTemplateObj,
                ref formatObj,
                ref encodingObj,
                ref visibleObj,
                ref openAndRepairObj,
                ref documentDirectionObj,
                ref noEncodingDialogObj,
                ref xMLTransformObj);

            var wordDoc = new WordDocument(doc, this);
            MemorizeActiveDocument(wordDoc);
            return wordDoc;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to open document '{filePath}'.", ex);
        }
    }

    public void Quit()
    {
        try
        {
            object saveChanges = missing;
            object originalFormat = missing;
            object routeDocument = missing;

            _application.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to quit Word application.", ex);
        }
    }

    public void Activate()
    {
        try
        {
            _application.Activate();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to activate Word application.", ex);
        }
    }

    public IOfficeFileDialog CreateFileDialog(MsoFileDialogType fileDialogType)
    {
        MsCore.FileDialog dialog = _application.FileDialog[(MsCore.MsoFileDialogType)fileDialogType];
        return new OfficeFileDialog(dialog);
    }

    public void RunMacro(string macroName)
    {
        if (string.IsNullOrEmpty(macroName))
            throw new ArgumentException("Macro name cannot be null or empty.", nameof(macroName));

        try
        {
            _application.Run(macroName);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to run macro '{macroName}'.", ex);
        }
    }

    /// <summary>
    /// 运行宏
    /// </summary>
    /// <param name="macroName">宏名称</param>
    /// <param name="args">宏参数</param>
    /// <returns>宏执行结果</returns>
    public object Run(string macroName, params object[] args)
    {
        if (_application == null || string.IsNullOrEmpty(macroName))
            return null;

        try
        {
            object result = _application.Run(macroName, args);
            return result;
        }
        catch
        {
            return null;
        }
    }

    public void Minimize()
    {
        try
        {
            _application.WindowState = MsWord.WdWindowState.wdWindowStateMinimize;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to minimize Word application.", ex);
        }
    }

    public void Maximize()
    {
        try
        {
            _application.WindowState = MsWord.WdWindowState.wdWindowStateMaximize;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to maximize Word application.", ex);
        }
    }

    public void Restore()
    {
        try
        {
            _application.WindowState = MsWord.WdWindowState.wdWindowStateNormal;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to restore Word application.", ex);
        }
    }

    public IEnumerable<string> GetRecentFiles(int count = 10)
    {
        if (count <= 0)
            throw new ArgumentOutOfRangeException(nameof(count), "Count must be greater than zero.");

        try
        {
            var recentFiles = new List<string>();
            var maxCount = Math.Min(count, _application.RecentFiles.Count);

            for (int i = 1; i <= maxCount; i++)
            {
                try
                {
                    recentFiles.Add(_application.RecentFiles[i].Name);
                }
                catch
                {
                    continue;
                }
            }

            return recentFiles;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get recent files.", ex);
        }
    }

    public void AddToRecentFiles(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("File not found.", filePath);

        try
        {
            _application.RecentFiles.Add(filePath);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add file to recent files.", ex);
        }
    }

    public void SetOption(string optionName, object value)
    {
        if (string.IsNullOrEmpty(optionName))
            throw new ArgumentException("Option name cannot be null or empty.", nameof(optionName));

        try
        {
            // Word Application 对象没有直接的 Options 属性设置方法
            // 这里需要根据具体的选项名称来实现
            throw new NotImplementedException($"Setting option '{optionName}' is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to set option '{optionName}'.", ex);
        }
    }

    public object GetOption(string optionName)
    {
        if (string.IsNullOrEmpty(optionName))
            throw new ArgumentException("Option name cannot be null or empty.", nameof(optionName));

        try
        {
            // Word Application 对象没有直接的 Options 属性获取方法
            // 这里需要根据具体的选项名称来实现
            throw new NotImplementedException($"Getting option '{optionName}' is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get option '{optionName}'.", ex);
        }
    }

    public string GetInstallPath()
    {
        try
        {
            // 获取 Word 安装路径的实现
            return System.IO.Path.GetDirectoryName(_application.Path);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get Word installation path.", ex);
        }
    }

    public string GetProductName()
    {
        try
        {
            return _application.Name;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get Word product name.", ex);
        }
    }

    public bool IsPrinting()
    {
        try
        {
            // Word Application 对象没有直接的打印状态属性
            // 这里返回 false 作为默认实现
            return false;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to check printing status.", ex);
        }
    }

    public void CancelPrintJobs()
    {
        try
        {
            // Word Application 对象没有直接的取消打印作业方法
            // 这里作为占位符实现
            throw new NotImplementedException("Cancel print jobs is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to cancel print jobs.", ex);
        }
    }

    public IWordSystemInfo GetSystemInfo()
    {
        try
        {
            return new WordSystemInfo
            {
                OSVersion = Environment.OSVersion.ToString(),
                TotalMemory = Environment.WorkingSet,
                AvailableMemory = 0, // .NET 中没有直接获取可用内存的方法
                ProcessorCount = Environment.ProcessorCount,
                SystemUpTime = DateTime.Now - TimeSpan.FromTicks(Environment.TickCount)
            };
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get system information.", ex);
        }
    }

    public void Refresh()
    {
        try
        {
            // Word Application 对象没有直接的刷新方法
            // 这里作为占位符实现
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh application.", ex);
        }
    }

    public IWordDocument BlankDocument()
    {
        try
        {
            var doc = _application.Documents.Add();
            var wordDoc = new WordDocument(doc, this);
            MemorizeActiveDocument(wordDoc);
            return wordDoc;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create blank document.", ex);
        }
    }


    internal WordApplication()
    {
        _application = new MsWord.Application();
        _applicationEvent = _application;
        InitializeApp();
        ConnectEvents();
    }

    internal WordApplication(MsWord.Application application)
    {
        _application = application ?? throw new ArgumentNullException(nameof(application));
        _applicationEvent = application;
        InitializeApp();
        ConnectEvents();
    }


    private void InitializeApp()
    {
        _application.DisplayAlerts = MsWord.WdAlertLevel.wdAlertsMessageBox;
        _disposedValue = false;
        _activeDocument = null;
    }

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

    private void MemorizeActiveDocument(IWordDocument document)
    {
        _activeDocument = document;
    }

    private static readonly object missing = System.Reflection.Missing.Value;

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放相关对象
            _selection?.Dispose();
            _documents?.Dispose();
            _windows?.Dispose();
            _wordTemplate?.Dispose();
            DisconnectEvents();
            if (_application != null)
            {
                try
                {
                    _application.DisplayAlerts = MsWord.WdAlertLevel.wdAlertsAll;

                    if (Visibility == WordAppVisibility.Hidden)
                    {
                        object saveChanges = MsWord.WdSaveOptions.wdDoNotSaveChanges;
                        object originalFormat = missing;
                        object routeDocument = missing;
                        _application.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
                    }
                    else
                    {
                        _application.Visible = true;
                    }

                    try { while (Marshal.ReleaseComObject(_application) > 0) ; } catch { }
                }
                catch
                {
                    // 忽略释放失败的情况
                }
            }

            GC.Collect();
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #region 事件处理方法
    private void OnDocumentOpen(MsWord.Document doc)
    {
        _documentOpen?.Invoke(new WordDocument(doc, this));
    }

    private void OnDocumentBeforeClose(MsWord.Document doc, ref bool cancel)
    {
        _documentBeforeClose?.Invoke(new WordDocument(doc, this), ref cancel);
    }

    private void OnDocumentBeforeSave(MsWord.Document doc, ref bool saveAsUI, ref bool cancel)
    {
        _documentBeforeSave?.Invoke(new WordDocument(doc, this), ref saveAsUI, ref cancel);
    }

    private void OnNewDocument(MsWord.Document doc)
    {
        _newDocument?.Invoke(new WordDocument(doc, this));
    }

    private void OnWindowActivate(MsWord.Document doc, MsWord.Window wnd)
    {
        _windowActivate?.Invoke(new WordDocument(doc, this), new WordWindow(wnd));
    }

    private void OnWindowDeactivate(MsWord.Document doc, MsWord.Window wnd)
    {
        _windowDeactivate?.Invoke(new WordDocument(doc, this), new WordWindow(wnd));
    }

    private void OnDocumentSync(MsWord.Document doc, MsCore.MsoSyncEventType syncEventType)
    {
        _documentSync?.Invoke(new WordDocument(doc, this), (MsoSyncEventType)syncEventType);
    }

    private void OnDocumentChange()
    {
        _documentChange?.Invoke();
    }

    private void OnMailMergeDataSourceLoad(MsWord.Document doc)
    {
        _mailMergeDataSourceLoad?.Invoke(new WordDocument(doc, this));
    }

    private void OnMailMergeDataSourceValidate(MsWord.Document doc, ref bool handled)
    {
        _mailMergeDataSourceValidate?.Invoke(new WordDocument(doc, this), ref handled);
    }

    private void OnWindowSelectionChange(MsWord.Selection sel)
    {
        _windowSelectionChange?.Invoke(new WordSelection(sel, _activeDocument));
    }

    private void OnWindowSize(MsWord.Document doc, MsWord.Window wnd)
    {
        _windowSize?.Invoke(new WordDocument(doc, this), new WordWindow(wnd));
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

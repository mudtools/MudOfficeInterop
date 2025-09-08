//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Imps;

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint Application 对象的二次封装实现类
/// 实现 IPowerPointApplication 接口
/// </summary>
internal class PowerPointApplication : IPowerPointApplication
{
    private MsPowerPoint.Application _application;
    private MsPowerPoint.EApplication_Event _application_Event;
    private bool _disposedValue = false;

    #region 事件字段
    private PresentationOpenEventHandler _presentationOpen;
    private NewPresentationEventHandler _newPresentation;
    private PresentationBeforeCloseEventHandler _presentationBeforeClose;
    private PresentationSaveEventHandler _presentationSave;
    private WindowActivateEventHandler _windowActivate;
    private WindowDeactivateEventHandler _windowDeactivate;
    private WindowSelectionChangeEventHandler _windowSelectionChange;
    private PresentationSyncEventHandler _presentationSync;
    private PresentationChangeEventHandler _presentationChange;
    private SlideShowBeginEventHandler _slideShowBegin;
    private SlideShowEndEventHandler _slideShowEnd;
    private SlideShowNextSlideEventHandler _slideShowNextSlide;
    #endregion

    /// <summary>
    /// 初始化 PowerPointApplication 实例
    /// </summary>
    /// <param name="application">要封装的 Microsoft.Office.Interop.PowerPoint.Application 对象</param>
    internal PowerPointApplication(MsPowerPoint.Application application)
    {
        _application = application ?? throw new ArgumentNullException(nameof(application));
        _application_Event = application;
        ConnectEvents();
    }

    internal PowerPointApplication()
    {
        _application = new MsPowerPoint.Application();
        _application_Event = _application;
        InitializeApp();
        ConnectEvents();
    }

    private void InitializeApp()
    {
        _application.DisplayAlerts = MsPowerPoint.PpAlertLevel.ppAlertsAll;
        _disposedValue = false;
    }

    #region 基础属性

    public string Name
    {
        get => _application.Name;
    }

    public string Version => _application.Version;

    public string Path => _application.Path;

    public bool Visible
    {
        get => _application.Visible == MsCore.MsoTriState.msoTrue;
        set => _application.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
    }

    public string? Build => _application?.Build;

    public bool IsActive
    {
        get
        {
            try
            {
                return _application.ActiveWindow != null;
            }
            catch
            {
                return false;
            }
        }
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
                _application.WindowState = (MsPowerPoint.PpWindowState)value;
        }
    }
    #endregion

    #region 应用程序状态
    public bool IsBusy => _application.IsSandboxed;
    #endregion

    #region 核心对象集合和属性
    public IPowerPointPresentations Presentations => new PowerPointPresentations(_application.Presentations);

    public IPowerPointDocumentWindows Windows => new PowerPointDocumentWindows(_application.Windows);

    public IPowerPointPresentation ActivePresentation => _application.ActivePresentation != null ? new PowerPointPresentation(_application.ActivePresentation) : null;

    public IPowerPointDocumentWindow ActiveWindow => _application.ActiveWindow != null ? new PowerPointDocumentWindow(_application.ActiveWindow) : null;

    public IPowerPointSlide ActiveSlide
    {
        get
        {
            try
            {
                var view = _application.ActiveWindow?.View;
                if (view?.Type == MsPowerPoint.PpViewType.ppViewNormal && view is MsPowerPoint.SlideShowView slideView)
                {
                    return slideView.Slide != null ? new PowerPointSlide(slideView.Slide) : null;
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
    }

    public IPowerPointSelection Selection => _application.ActiveWindow?.Selection != null ? new PowerPointSelection(_application.ActiveWindow.Selection) : null;
    public IPowerPointView ActiveView => _application.ActiveWindow?.View != null ? new PowerPointView(_application.ActiveWindow.View, this.ActivePresentation) : null;
    #endregion

    #region 环境和设置
    public IOfficeLanguageSettings LanguageSettings => new OfficeLanguageSettings(_application.LanguageSettings);

    public IOfficeCommandBars CommandBars => new OfficeCommandBars(_application.CommandBars);

    public float Left
    {
        get => _application.Left;
        set => _application.Left = value;
    }

    public float Top
    {
        get => _application.Top;
        set => _application.Top = value;
    }

    public float Width
    {
        get => _application.Width;
        set => _application.Width = value;
    }

    public float Height
    {
        get => _application.Height;
        set => _application.Height = value;
    }
    #endregion

    #region 操作方法
    public void Activate()
    {
        _application.Activate();
    }

    public void Quit()
    {
        _application.Quit();
    }

    public void RunCommand(string commandId)
    {
        try
        {
            _application.CommandBars.ExecuteMso(commandId);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error executing command '{commandId}': {ex.Message}");
        }
    }

    public object Run(string macroName, params object[] args)
    {
        try
        {
            return _application.Run(macroName, args);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error running macro '{macroName}': {ex.Message}");
            return null;
        }
    }



    public void SaveAll()
    {
        try
        {
            var presentations = _application.Presentations;
            for (int i = 1; i <= presentations.Count; i++)
            {
                presentations[i].Save();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error saving all presentations: {ex.Message}");
        }
    }
    #endregion

    #region 文件操作

    public IPowerPointPresentation BlankDocument()
    {
        try
        {
            var doc = _application.Presentations.Add();
            var presentation = new PowerPointPresentation(doc);
            return presentation;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create blank document.", ex);
        }
    }

    public IPowerPointPresentation OpenPresentation(string filename, bool readOnly = false, bool untitled = false, bool withWindow = true)
    {
        try
        {
            var presentation = _application.Presentations.Open(
                filename,
                readOnly ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                untitled ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                withWindow ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse
            );
            return new PowerPointPresentation(presentation);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error opening presentation '{filename}': {ex.Message}");
            return null;
        }
    }

    public IPowerPointPresentation AddPresentation(bool withWindow = true)
    {
        try
        {
            var presentation = _application.Presentations.Add(
                withWindow ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse
            );
            return new PowerPointPresentation(presentation);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error adding new presentation: {ex.Message}");
            return null;
        }
    }
    #endregion

    #region UI 和交互
    public IOfficeFileDialog CreateFileDialog(MsoFileDialogType fileDialogType)
    {
        MsCore.FileDialog dialog = _application.FileDialog[(MsCore.MsoFileDialogType)fileDialogType];
        return new OfficeFileDialog(dialog);
    }

    #endregion

    #region 事件处理

    private void ConnectEvents()
    {
        _application_Event.NewPresentation += OnNewPresentation;
        _application_Event.PresentationOpen += OnPresentationOpen;
        _application_Event.PresentationClose += OnPresentationBeforeClose;
        _application_Event.PresentationSave += OnPresentationSave;
        _application_Event.WindowActivate += OnWindowActivate;
        _application_Event.WindowDeactivate += OnWindowDeactivate;
        _application_Event.WindowSelectionChange += OnWindowSelectionChange;
        _application_Event.PresentationSync += OnPresentationSync;
        _application_Event.PresentationNewSlide += OnPresentationChange;
        _application_Event.SlideShowBegin += OnSlideShowBegin;
        _application_Event.SlideShowEnd += OnSlideShowEnd;
        _application_Event.SlideShowNextSlide += OnSlideShowNextSlide;
    }


    private void DisconnectEvents()
    {
        if (_application_Event == null) return;

        _application_Event.NewPresentation -= OnNewPresentation;
        _application.PresentationOpen -= OnPresentationOpen;
        _application.PresentationClose -= OnPresentationBeforeClose;
        _application.PresentationSave -= OnPresentationSave;
        _application.WindowActivate -= OnWindowActivate;
        _application.WindowDeactivate -= OnWindowDeactivate;
        _application.WindowSelectionChange -= OnWindowSelectionChange;
        _application.PresentationSync -= OnPresentationSync;
        _application.PresentationNewSlide -= OnPresentationChange;
        _application.SlideShowBegin -= OnSlideShowBegin;
        _application.SlideShowEnd -= OnSlideShowEnd;
        _application.SlideShowNextSlide -= OnSlideShowNextSlide;
        _application_Event = null;
    }

    private void OnNewPresentation(MsPowerPoint.Presentation Pres)
    {
        _newPresentation?.Invoke(new PowerPointPresentation(Pres));
    }

    private void OnPresentationOpen(MsPowerPoint.Presentation Pres)
    {
        _presentationOpen?.Invoke(new PowerPointPresentation(Pres));
    }

    private void OnPresentationBeforeClose(MsPowerPoint.Presentation Pres)
    {
        _presentationBeforeClose?.Invoke(new PowerPointPresentation(Pres));
    }

    private void OnPresentationSave(MsPowerPoint.Presentation Pres)
    {
        _presentationSave?.Invoke(new PowerPointPresentation(Pres));
    }

    private void OnWindowActivate(MsPowerPoint.Presentation Pres, MsPowerPoint.DocumentWindow Wn)
    {
        _windowActivate?.Invoke(new PowerPointPresentation(Pres), new PowerPointDocumentWindow(Wn));
    }

    private void OnWindowDeactivate(MsPowerPoint.Presentation Pres, MsPowerPoint.DocumentWindow Wn)
    {
        _windowDeactivate?.Invoke(new PowerPointPresentation(Pres), new PowerPointDocumentWindow(Wn));
    }

    private void OnWindowSelectionChange(MsPowerPoint.Selection Sel)
    {
        _windowSelectionChange?.Invoke(new PowerPointSelection(Sel));
    }

    private void OnPresentationSync(MsPowerPoint.Presentation Pres, Microsoft.Office.Core.MsoSyncEventType Type)
    {
        _presentationSync?.Invoke(new PowerPointPresentation(Pres), (MsoSyncEventType)Type);
    }

    private void OnPresentationChange(MsPowerPoint.Slide Sld)
    {
        _presentationChange?.Invoke();
    }

    private void OnSlideShowBegin(MsPowerPoint.SlideShowWindow Wn)
    {
        _slideShowBegin?.Invoke(new PowerPointSlideShowWindow(Wn));
    }

    private void OnSlideShowEnd(MsPowerPoint.Presentation Wn)
    {
        _slideShowEnd?.Invoke(new PowerPointPresentation(Wn));
    }

    private void OnSlideShowNextSlide(MsPowerPoint.SlideShowWindow Wn)
    {
        _slideShowNextSlide?.Invoke(new PowerPointSlideShowWindow(Wn));
    }

    public event NewPresentationEventHandler NewPresentation
    {
        add { _newPresentation += value; }
        remove { _newPresentation -= value; }
    }


    public event PresentationOpenEventHandler PresentationOpen
    {
        add { _presentationOpen += value; }
        remove { _presentationOpen -= value; }
    }

    public event PresentationBeforeCloseEventHandler PresentationBeforeClose
    {
        add { _presentationBeforeClose += value; }
        remove { _presentationBeforeClose -= value; }
    }

    public event PresentationSaveEventHandler PresentationSave
    {
        add { _presentationSave += value; }
        remove { _presentationSave -= value; }
    }

    public event WindowActivateEventHandler WindowActivate
    {
        add { _windowActivate += value; }
        remove { _windowActivate -= value; }
    }

    public event WindowDeactivateEventHandler WindowDeactivate
    {
        add { _windowDeactivate += value; }
        remove { _windowDeactivate -= value; }
    }

    public event WindowSelectionChangeEventHandler WindowSelectionChange
    {
        add { _windowSelectionChange += value; }
        remove { _windowSelectionChange -= value; }
    }

    public event PresentationSyncEventHandler PresentationSync
    {
        add { _presentationSync += value; }
        remove { _presentationSync -= value; }
    }

    public event PresentationChangeEventHandler PresentationChange
    {
        add { _presentationChange += value; }
        remove { _presentationChange -= value; }
    }

    public event SlideShowBeginEventHandler SlideShowBegin
    {
        add { _slideShowBegin += value; }
        remove { _slideShowBegin -= value; }
    }

    public event SlideShowEndEventHandler SlideShowEnd
    {
        add { _slideShowEnd += value; }
        remove { _slideShowEnd -= value; }
    }

    public event SlideShowNextSlideEventHandler SlideShowNextSlide
    {
        add { _slideShowNextSlide += value; }
        remove { _slideShowNextSlide -= value; }
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (_application != null)
            {
                try
                {
                    DisconnectEvents();
                    Marshal.FinalReleaseComObject(_application);
                }
                catch
                {
                    // 忽略释放过程中可能发生的异常
                }
                _application = null;
            }

            _disposedValue = true;
        }
    }

    ~PowerPointApplication()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Vbe;
using MudTools.OfficeInterop.Vbe.Imp;

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Workbook 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Workbook 对象的安全访问和资源管理
/// </summary>
internal partial class ExcelWorkbook : IExcelWorkbook
{
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelWorkbook));
    /// <summary>
    /// 底层的 COM Workbook 对象
    /// </summary>
    internal MsExcel.Workbook _workbook;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelWorkbook 实例
    /// </summary>
    /// <param name="workbook">底层的 COM Workbook 对象</param>
    internal ExcelWorkbook(MsExcel.Workbook workbook)
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        _workbookEvents_Event = workbook;
        InitializeEvents();
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                _excelConnections?.Dispose();
                _windows?.Dispose();
                _vbeVBProject?.Dispose();
                _worksheets?.Dispose();
                _sheets?.Dispose();
                _modules?.Dispose();
                _names?.Dispose();
                _styles?.Dispose();
                _slicerCaches?.Dispose();
                _activeSlicer?.Dispose();
                _charts?.Dispose();
                _activeWorksheet?.Dispose();
                DisConnectEvent();
                // 释放底层COM对象
                if (_workbook != null)
                    Marshal.ReleaseComObject(_workbook);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
        }
        _slicerCaches = null;
        _activeSlicer = null;
        _workbookEvents_Event = null;
        _worksheets = null;
        _sheets = null;
        _modules = null;
        _names = null;
        _styles = null;
        _charts = null;
        _excelConnections = null;
        _activeWorksheet = null;
        _windows = null;
        _vbeVBProject = null;
        _workbook = null;
        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    public bool HasPassword => _workbook.HasPassword;

    /// <summary>
    /// 获取或设置工作簿的名称
    /// </summary>
    public string Name => _workbook?.Name;

    /// <summary>
    /// 获取工作簿的完整路径
    /// </summary>
    public string FullName => _workbook?.FullName;

    /// <summary>
    /// 获取工作簿的路径
    /// </summary>
    public string Path => _workbook?.Path;


    public bool MultiUserEditing => _workbook.MultiUserEditing;

    /// <summary>
    /// 获取或设置工作簿是否已保存
    /// </summary>
    public bool Saved
    {
        get => _workbook != null && _workbook.Saved;
        set
        {
            if (_workbook != null)
                _workbook.Saved = value;
        }
    }

    public string? Keywords
    {
        get => _workbook?.Keywords;
        set
        {
            if (_workbook != null)
                _workbook.Keywords = value;
        }
    }

    public string? OnSave
    {
        get => _workbook?.OnSave;
        set
        {
            if (_workbook != null)
                _workbook.OnSave = value;
        }
    }

    public string? OnSheetActivate
    {
        get => _workbook?.OnSheetActivate;
        set
        {
            if (_workbook != null)
                _workbook.OnSheetActivate = value;
        }
    }

    public string? OnSheetDeactivate
    {
        get => _workbook?.OnSheetDeactivate;
        set
        {
            if (_workbook != null)
                _workbook.OnSheetDeactivate = value;
        }
    }

    public string? Subject
    {
        get => _workbook?.Subject;
        set
        {
            if (_workbook != null)
                _workbook.Subject = value;
        }
    }

    public bool IsAddin
    {
        get => _workbook != null && _workbook.IsAddin;
    }

    public bool ProtectStructure
    {
        get => _workbook != null && _workbook.ProtectStructure;
    }

    public bool ProtectWindows
    {
        get => _workbook != null && _workbook.ProtectWindows;
    }



    public bool PersonalViewListSettings
    {
        get => _workbook != null && _workbook.PersonalViewListSettings;
    }

    public bool PersonalViewPrintSettings
    {
        get => _workbook != null && _workbook.PersonalViewPrintSettings;
    }

    public bool PrecisionAsDisplayed
    {
        get => _workbook != null && _workbook.PrecisionAsDisplayed;
        set
        {
            if (_workbook != null)
                _workbook.PrecisionAsDisplayed = value;
        }
    }

    public bool HasVBProject
    {
        get => _workbook != null && _workbook.HasVBProject;
    }

    public XlDisplayDrawingObjects DisplayDrawingObjects
    {
        get => _workbook.DisplayDrawingObjects.EnumConvert(XlDisplayDrawingObjects.xlHide);
        set
        {
            if (_workbook != null)
                _workbook.DisplayDrawingObjects = value.EnumConvert(MsExcel.XlDisplayDrawingObjects.xlHide);
        }
    }

    public XlFileFormat FileFormat
    {
        get => _workbook.FileFormat.EnumConvert(XlFileFormat.xlWorkbookDefault);
    }

    /// <summary>
    /// 获取工作簿是否受保护
    /// </summary>
    public bool IsProtected => _workbook != null && _workbook.ProtectStructure;

    /// <summary>
    /// 获取工作簿的只读状态
    /// </summary>
    public bool ReadOnly => _workbook != null && _workbook.ReadOnly;

    private IExcelConnections _excelConnections;
    public IExcelConnections? Connections
    {
        get
        {
            if (_excelConnections == null)
                return null;
            if (_excelConnections != null)
                return _excelConnections;
            _excelConnections = new ExcelConnections(_workbook.Connections);
            return _excelConnections;
        }
    }

    private IVbeVBProject _vbeVBProject;

    public IVbeVBProject? VBProject
    {
        get
        {
            if (_workbook == null)
                return null;
            if (_vbeVBProject != null)
                return _vbeVBProject;
            _vbeVBProject = new VbeVBProject(_workbook.VBProject);
            return _vbeVBProject;
        }
    }


    private IExcelWindows _windows;
    public IExcelWindows? Windows
    {
        get
        {
            if (_workbook == null)
                return null;
            if (_windows != null)
                return _windows;

            _windows = new ExcelWindows(_workbook.Windows, Application);
            return _windows;
        }
    }

    /// <summary>
    /// 获取工作簿的修改时间
    /// </summary>
    public DateTime ModifiedTime
    {
        get
        {
            try
            {
                if (!string.IsNullOrEmpty(FullName) && System.IO.File.Exists(FullName))
                {
                    return System.IO.File.GetLastWriteTime(FullName);
                }
            }
            catch
            {
                // 忽略异常
            }
            return DateTime.Now;
        }
    }

    /// <summary>
    /// 获取工作簿的创建时间
    /// </summary>
    public DateTime CreatedTime
    {
        get
        {
            // 注意：COM对象通常不直接提供创建时间属性
            // 需要通过文件系统获取
            try
            {
                if (!string.IsNullOrEmpty(FullName) && System.IO.File.Exists(FullName))
                {
                    return System.IO.File.GetCreationTime(FullName);
                }
            }
            catch
            {
                // 忽略异常
            }
            return DateTime.Now;
        }
    }

    /// <summary>
    /// 获取工作簿的文件大小（字节）
    /// </summary>
    public long FileSize
    {
        get
        {
            // 注意：COM对象通常不直接提供文件大小属性
            // 需要通过文件系统获取
            try
            {
                if (!string.IsNullOrEmpty(FullName) && System.IO.File.Exists(FullName))
                {
                    return new System.IO.FileInfo(FullName).Length;
                }
            }
            catch
            {
                // 忽略异常
            }
            return 0;
        }
    }

    /// <summary>
    /// 获取工作簿所在的父对象
    /// </summary>
    public object? Parent
    {
        get
        {
            if (_workbook.Parent == null)
                return null;
            if (_workbook.Parent is MsExcel.Application app)
                return new ExcelApplication(app);

            if (_workbook.Parent is MsExcel.Window win)
                return new ExcelWindow(win);

            if (_workbook.Parent is MsExcel.Windows wins)
                return new ExcelWindows(wins, Application);

            return _workbook.Parent;
        }
    }

    /// <summary>
    /// 获取工作簿所在的Application对象
    /// </summary>
    public IExcelApplication Application
    {
        get
        {
            return _workbook?.Application is MsExcel.Application application ? new ExcelApplication(application) : null;
        }
    }

    /// <summary>
    /// 获取工作簿的编码名称
    /// </summary>
    public string CodeName => _workbook?.CodeName;

    #endregion

    #region 工作表管理   

    /// <summary>
    /// 工作表集合缓存
    /// </summary>
    private IExcelWorksheets? _worksheets;

    /// <summary>
    /// 获取工作簿中的所有工作表集合
    /// </summary>
    public IExcelWorksheets Worksheets => _worksheets ??= new ExcelWorksheets(_workbook?.Worksheets);

    /// <summary>
    /// 工作表集合缓存（所有类型）
    /// </summary>
    private IExcelSheets? _sheets;

    private IExcelSheets? _modules;

    /// <summary>
    /// 获取工作簿中的所有工作表集合（包括图表工作表等）
    /// </summary>
    public IExcelSheets? Sheets => _sheets ??= new ExcelSheets(_workbook?.Sheets);

    public IExcelSheets? Modules => _modules ??= new ExcelSheets(_workbook?.Modules);

    /// <summary>
    /// 获取工作簿中的工作表数量
    /// </summary>
    public int WorksheetCount => _workbook?.Worksheets?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的工作表
    /// </summary>
    /// <param name="index">工作表索引</param>
    /// <returns>工作表对象</returns>
    public IExcelWorksheet? GetWorksheet(int index)
    {
        if (_workbook?.Worksheets == null || index < 1 || index > WorksheetCount)
            return null;

        try
        {
            var worksheet = _workbook.Worksheets[index] as MsExcel.Worksheet;
            return worksheet != null ? new ExcelWorksheet(worksheet) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 获取指定名称的工作表
    /// </summary>
    /// <param name="name">工作表名称</param>
    /// <returns>工作表对象</returns>
    public IExcelWorksheet? GetWorksheet(string name)
    {
        if (_workbook?.Worksheets == null || string.IsNullOrEmpty(name))
            return null;

        try
        {
            var worksheet = _workbook.Worksheets[name] as MsExcel.Worksheet;
            return worksheet != null ? new ExcelWorksheet(worksheet) : null;
        }
        catch
        {
            return null;
        }
    }
    public IExcelWindow? NewWindow()
    {
        if (_workbook == null)
            return null;
        return new ExcelWindow(_workbook.NewWindow());
    }

    /// <summary>
    /// 添加新的工作表
    /// </summary>
    /// <param name="before">添加到指定工作表之前</param>
    /// <param name="after">添加到指定工作表之后</param>
    /// <param name="count">添加的工作表数量</param>
    /// <param name="type">工作表类型</param>
    /// <returns>新创建的工作表对象</returns>
    public IExcelWorksheet? AddWorksheet(IExcelWorksheet before = null, IExcelWorksheet after = null,
                                      int count = 1, int type = 0)
    {
        if (_workbook?.Worksheets == null)
            return null;

        try
        {
            ExcelWorksheet? beforeSheet = before as ExcelWorksheet;
            ExcelWorksheet? afterSheet = after as ExcelWorksheet;

            return _workbook.Worksheets.Add(
                                    beforeSheet?.Worksheet,
                                    afterSheet?.Worksheet,
                                    count,
                (MsExcel.XlSheetType)type
            ) is MsExcel.Worksheet worksheet ? new ExcelWorksheet(worksheet) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 删除工作表
    /// </summary>
    /// <param name="worksheet">要删除的工作表</param>
    public void DeleteWorksheet(IExcelWorksheet worksheet)
    {
        if (_workbook?.Worksheets == null || worksheet == null)
            return;

        try
        {
            worksheet.Delete();
        }
        catch
        {
            // 忽略删除过程中的异常
        }
    }

    public void ExclusiveAccess()
    {
        _workbook?.ExclusiveAccess();
    }

    /// <summary>
    /// 活动工作表缓存
    /// </summary>
    private IExcelComSheet? _activeWorksheet;

    /// <summary>
    /// 获取活动工作表
    /// </summary>
    public IExcelComSheet? ActiveSheet
    {
        get
        {
            if (_activeWorksheet != null)
                return _activeWorksheet;

            var sheet = _workbook.ActiveSheet;
            if (sheet != null && sheet is MsExcel.Worksheet worksheet)
                _activeWorksheet = new ExcelWorksheet(worksheet);
            if (sheet != null && sheet is MsExcel.Chart chart)
                _activeWorksheet = new ExcelChart(chart);
            return _activeWorksheet;
        }
    }

    public IExcelWorksheet? ActiveSheetWrap
    {
        get
        {
            if (ActiveSheet is IExcelWorksheet worksheet)
                return worksheet;
            return null;
        }
    }


    #endregion

    #region 保护和安全
    public void ChangeFileAccess(XlFileAccess Mode, string? WritePassword, bool? Notify)
    {
        _workbook?.ChangeFileAccess(Mode.EnumConvert(MsExcel.XlFileAccess.xlReadOnly), WritePassword.ComArgsVal(), Notify.ComArgsVal());
    }

    public void DeleteNumberFormat(string NumberFormatName)
    {
        _workbook?.DeleteNumberFormat(NumberFormatName);
    }



    /// <summary>
    /// 保护工作簿结构和窗口
    /// </summary>
    /// <param name="password">保护密码</param>
    /// <param name="structure">是否保护结构</param>
    /// <param name="windows">是否保护窗口</param>
    public void Protect(string password = "", bool structure = true, bool windows = false)
    {
        _workbook?.Protect(password, structure, windows);
    }

    /// <summary>
    /// 取消保护工作簿
    /// </summary>
    /// <param name="password">保护密码</param>
    public void Unprotect(string password = "")
    {
        _workbook?.Unprotect(password);
    }

    /// <summary>
    /// 保护工作簿中的所有工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    public void ProtectAllWorksheets(string password = "")
    {
        var worksheets = Sheets;
        if (worksheets != null)
        {
            worksheets.ProtectAll(password);
        }
    }

    /// <summary>
    /// 取消保护工作簿中的所有工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    public void UnprotectAllWorksheets(string password = "")
    {
        // 通过Worksheets集合实现
        var worksheets = Sheets;
        if (worksheets != null)
        {
            worksheets.UnprotectAll(password);
        }
    }

    #endregion

    #region 操作方法

    public void OpenLinks(string? name, bool? readOnly, XlLink? type)
    {
        _workbook?.OpenLinks(name, readOnly.ComArgsVal(), type.ComArgsConvert(d => d.EnumConvert(MsExcel.XlLink.xlExcelLinks)));
    }
    public object? LinkInfo(string? name, XlLinkInfo linkInfo, XlLinkInfoType? linkInfoType = null)
    {

        return _workbook?.LinkInfo(name != null ? Name : "",
                                    linkInfo.EnumConvert(MsExcel.XlLinkInfo.xlUpdateState),
                                    linkInfoType.ComArgsConvert(d => d.EnumConvert(MsExcel.XlLinkInfoType.xlLinkInfoOLELinks)));
    }


    public void ExportAsFixedFormat(
        XlFixedFormatType Type,
        string Filename,
        object? Quality = null,
        object? IncludeDocProperties = null,
        object? IgnorePrintAreas = null,
        object? From = null,
        object? To = null,
        object? OpenAfterPublish = null,
        object? FixedFormatExtClassPtr = null)
    {
        object? missing = System.Type.Missing;
        _workbook.ExportAsFixedFormat((MsExcel.XlFixedFormatType)(int)Type,
            Filename, Quality ?? missing, IncludeDocProperties ?? missing, IgnorePrintAreas ?? missing,
            From ?? missing, To ?? missing, OpenAfterPublish ?? missing, FixedFormatExtClassPtr ?? missing
            );
    }


    /// <summary>
    /// 保存工作簿
    /// </summary>
    public void Save()
    {
        _workbook?.Save();
    }

    /// <summary>
    /// 另存为工作簿
    /// </summary>
    /// <param name="filename">文件路径</param>
    /// <param name="fileFormat">文件格式</param>
    /// <param name="password">打开密码</param>
    /// <param name="writeResPassword">写入密码</param>
    /// <param name="readOnlyRecommended">是否建议只读</param>
    /// <param name="createBackup">是否创建备份</param>
    /// <param name="accessMode">访问模式</param>
    /// <param name="conflictResolution">冲突解决方式</param>
    /// <param name="addToMru">是否添加到最近使用文件</param>
    /// <param name="local">是否本地格式</param>
    public void SaveAs(string filename, XlFileFormat fileFormat = XlFileFormat.xlWorkbookDefault, string? password = null,
                      string? writeResPassword = null, bool? readOnlyRecommended = false, bool? createBackup = false,
                      XlSaveAsAccessMode accessMode = XlSaveAsAccessMode.xlNoChange,
                      XlSaveConflictResolution? conflictResolution = XlSaveConflictResolution.xlLocalSessionChanges,
                      bool? addToMru = true, bool? local = false)
    {
        if (_workbook == null || string.IsNullOrEmpty(filename))
            return;

        try
        {
            _workbook.SaveAs(
                    filename, (MsExcel.XlFileFormat)(int)fileFormat,
                    password.ComArgsVal(), writeResPassword.ComArgsVal(),
                    readOnlyRecommended.ComArgsVal(),
                    createBackup.ComArgsVal(),
                    (MsExcel.XlSaveAsAccessMode)(int)accessMode,
                    conflictResolution.ComArgsVal(),
                    addToMru.ComArgsVal(), local.ComArgsVal()
            );
        }
        catch (COMException ce)
        {
            log.Error($"保存文件{filename}失败:{ce.Message}", ce);
        }
        catch (Exception ex)
        {
            log.Error($"保存文件{filename}失败:{ex.Message}", ex);
        }
    }

    /// <summary>
    /// 关闭工作簿
    /// </summary>
    /// <param name="saveChanges">是否保存更改</param>
    /// <param name="filename">文件路径</param>
    /// <param name="routeWorkbook">是否发送路由</param>
    public void Close(bool saveChanges = true, string filename = "", bool routeWorkbook = false)
    {
        _workbook?.Close(saveChanges, filename, routeWorkbook);
    }

    /// <summary>
    /// 激活工作簿
    /// </summary>
    public void Activate()
    {
        _workbook?.Activate();
    }

    /// <summary>
    /// 选择工作簿
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    public void Select(bool replace = true)
    {
        // 工作簿通常不直接选择，而是激活
        Activate();
    }

    /// <summary>
    /// 打印工作簿
    /// </summary>
    /// <param name="preview">是否打印预览</param>
    public void PrintOut(bool preview = false)
    {
        if (_workbook == null) return;

        if (preview)
        {
            _workbook.PrintPreview();
        }
        else
        {
            _workbook.PrintOut(
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing
            );
        }
    }

    /// <summary>
    /// 发送工作簿
    /// </summary>
    /// <param name="recipients">收件人</param>
    /// <param name="subject">主题</param>
    /// <param name="returnReceipt">是否要求回执</param>
    public void SendMail(string recipients, string subject = "", bool returnReceipt = false)
    {
        if (_workbook == null || string.IsNullOrEmpty(recipients))
            return;

        try
        {
            _workbook.SendMail(recipients, subject, returnReceipt);
        }
        catch (Exception x)
        {
            log.Error($"发送文件失败:{x.Message}", x);
        }
    }

    #endregion

    #region 高级功能
    /// <summary>
    /// 计算所有工作表
    /// </summary>
    public void CalculateAll()
    {
        if (_worksheets == null) return;

        try
        {
            for (int i = 1; i <= _worksheets.Count; i++)
            {
                if (_worksheets[i] != null
                      && _worksheets[i] is IExcelWorksheet worksheet)
                    worksheet?.Calculate();
            }
        }
        catch (Exception x)
        {
            log.Error($"计算工作表失败:{x.Message}", x);
        }
    }

    public void Reply()
    {
        try
        {
            _workbook?.Reply();
        }
        catch (Exception x)
        {
            log.Error($"回复邮件失败:{x.Message}", x);
        }
    }

    public void ReplyAll()
    {
        try
        {
            _workbook?.ReplyAll();
        }
        catch (Exception x)
        {
            log.Error($"回复邮件失败:{x.Message}", x);
        }
    }

    public void RemoveUser(int index)
    {
        try
        {
            _workbook?.RemoveUser(index);
        }
        catch (Exception x)
        {
            log.Error($"移除用户失败:{x.Message}", x);
        }
    }

    public void Route()
    {
        try
        {
            _workbook?.Route();
        }
        catch (Exception x)
        {
            log.Error($"发送文件失败:{x.Message}", x);
        }
    }

    public bool Routed
    {
        get
        {
            return _workbook?.Routed ?? false;
        }
    }


    /// <summary>
    /// 刷新工作簿
    /// </summary>
    public void RefreshAll()
    {
        try
        {
            _workbook?.RefreshAll();
        }
        catch (Exception x)
        {
            log.Error($"刷新工作簿失败:{x.Message}", x);
        }
    }

    /// <summary>
    /// 清除工作簿中的所有内容
    /// </summary>
    public void ClearAll()
    {
        var worksheets = Sheets;
        if (worksheets != null)
        {
            try
            {
                for (int i = 1; i <= worksheets.Count; i++)
                {
                    if (worksheets[i] != null
                        && worksheets[i] is IExcelWorksheet worksheet)
                        worksheet.ClearAll();
                }
            }
            catch (Exception x)
            {
                log.Error($"清除工作簿内容失败:{x.Message}", x);

            }
        }
    }

    /// <summary>
    /// 名称集合缓存
    /// </summary>
    private IExcelNames? _names;

    /// <summary>
    /// 获取工作簿的名称集合
    /// </summary>
    public IExcelNames Names => _names ??= new ExcelNames(_workbook?.Names);

    /// <summary>
    /// 样式集合缓存
    /// </summary>
    private IExcelStyles? _styles;

    /// <summary>
    /// 获取工作簿的样式集合
    /// </summary>
    public IExcelStyles Styles => _styles ??= new ExcelStyles(_workbook?.Styles);

    private IExcelSlicerCaches _slicerCaches;

    public IExcelSlicerCaches SlicerCaches => _slicerCaches ??= new ExcelSlicerCaches(_workbook?.SlicerCaches);

    private IExcelSlicer _activeSlicer;

    public IExcelSlicer ActiveSlicer => _activeSlicer ??= new ExcelSlicer(_workbook?.ActiveSlicer);

    /// <summary>
    /// 图表集合缓存
    /// </summary>
    private IExcelSheets? _charts;

    /// <summary>
    /// 获取工作簿的图表集合
    /// </summary>
    public IExcelSheets Charts => _charts ??= new ExcelSheets(_workbook?.Charts);

    /// <summary>
    /// 获取工作簿的透视表缓存集合
    /// </summary>
    public IExcelPivotCaches? PivotCaches()
    {
        try
        {
            var caches = _workbook?.PivotCaches();
            if (caches != null)
                return new ExcelPivotCaches(caches);
            return null;
        }
        catch (Exception x)
        {
            log.Error($"获取透视表缓存集合失败:{x.Message}", x);
            return null;
        }
    }

    #endregion

    #region 属性设置
    /// <summary>
    /// 获取或设置是否显示滚动条
    /// </summary>
    public bool DisplayScrollBars
    {
        get => _workbook?.Application?.DisplayScrollBars ?? false;
        set
        {
            if (_workbook?.Application != null)
                _workbook.Application.DisplayScrollBars = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示公式栏
    /// </summary>
    public bool DisplayFormulaBar
    {
        get => _workbook?.Application?.DisplayFormulaBar ?? false;
        set
        {
            if (_workbook?.Application != null)
                _workbook.Application.DisplayFormulaBar = value;
        }
    }

    /// <summary>
    /// 获取或设置窗口状态
    /// </summary>
    public XlWindowState WindowState
    {
        get => (XlWindowState)(_workbook?.Windows[1]?.WindowState.EnumConvert(XlWindowState.xlNormal));
        set
        {
            if (_workbook?.Windows[1] != null)
                _workbook.Windows[1].WindowState = value.EnumConvert(MsExcel.XlWindowState.xlNormal);
        }
    }

    /// <summary>
    /// 获取或设置窗口高度
    /// </summary>
    public double Height
    {
        get => _workbook?.Windows[1]?.Height ?? 0;
        set
        {
            if (_workbook?.Windows[1] != null)
                _workbook.Windows[1].Height = value;
        }
    }

    /// <summary>
    /// 获取或设置窗口宽度
    /// </summary>
    public double Width
    {
        get => _workbook?.Windows[1]?.Width ?? 0;
        set
        {
            if (_workbook?.Windows[1] != null)
                _workbook.Windows[1].Width = value;
        }
    }

    /// <summary>
    /// 获取或设置窗口左边距
    /// </summary>
    public double Left
    {
        get => _workbook?.Windows[1]?.Left ?? 0;
        set
        {
            if (_workbook?.Windows[1] != null)
                _workbook.Windows[1].Left = value;
        }
    }

    /// <summary>
    /// 获取或设置窗口顶边距
    /// </summary>
    public double Top
    {
        get => _workbook?.Windows[1]?.Top ?? 0;
        set
        {
            if (_workbook?.Windows[1] != null)
                _workbook.Windows[1].Top = value;
        }
    }

    /// <summary>
    /// 获取或设置是否启用事件
    /// </summary>
    public bool EnableEvents
    {
        get => _workbook?.Application?.EnableEvents ?? false;
        set
        {
            if (_workbook?.Application != null)
                _workbook.Application.EnableEvents = value;
        }
    }


    /// <summary>
    /// 获取或设置是否启用多线程计算
    /// </summary>
    public bool MultiThreadedCalculation
    {
        get => _workbook?.Application?.MultiThreadedCalculation?.Enabled ?? false;
        set
        {
            if (_workbook?.Application?.MultiThreadedCalculation != null)
                _workbook.Application.MultiThreadedCalculation.Enabled = value;
        }
    }

    /// <summary>
    /// 获取或设置计算模式
    /// </summary>
    public XlCalculation Calculation
    {
        get => _workbook.Application.Calculation.EnumConvert(XlCalculation.xlCalculationAutomatic);
        set
        {
            if (_workbook?.Application != null)
                _workbook.Application.Calculation = value.EnumConvert(MsExcel.XlCalculation.xlCalculationAutomatic);
        }
    }

    #endregion
}

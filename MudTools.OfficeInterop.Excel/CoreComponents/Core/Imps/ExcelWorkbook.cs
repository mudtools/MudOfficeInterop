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
                _names?.Dispose();
                _styles?.Dispose();
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
        _workbookEvents_Event = null;
        _worksheets = null;
        _sheets = null;
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
    public string Name => _workbook?.Name?.ToString();

    /// <summary>
    /// 获取工作簿的完整路径
    /// </summary>
    public string FullName => _workbook?.FullName?.ToString();

    /// <summary>
    /// 获取工作簿的路径
    /// </summary>
    public string Path => _workbook?.Path?.ToString();


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

    public bool ProtectStructure
    {
        get => _workbook != null && _workbook.ProtectStructure;
    }

    public XlDisplayDrawingObjects DisplayDrawingObjects
    {
        get => (XlDisplayDrawingObjects)_workbook.DisplayDrawingObjects;
        set
        {
            if (_workbook != null)
                _workbook.DisplayDrawingObjects = (MsExcel.XlDisplayDrawingObjects)value;
        }
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
            MsExcel.Application? application = _workbook?.Application as MsExcel.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    /// <summary>
    /// 获取工作簿的编码名称
    /// </summary>
    public string CodeName => _workbook?.CodeName?.ToString();

    #endregion

    #region 工作表管理

    /// <summary>
    /// 工作表集合缓存
    /// </summary>
    private IExcelSheets? _worksheets;

    /// <summary>
    /// 获取工作簿中的所有工作表集合
    /// </summary>
    public IExcelSheets Worksheets => _worksheets ??= new ExcelSheets(_workbook?.Worksheets);

    /// <summary>
    /// 工作表集合缓存（所有类型）
    /// </summary>
    private IExcelSheets? _sheets;

    /// <summary>
    /// 获取工作簿中的所有工作表集合（包括图表工作表等）
    /// </summary>
    public IExcelSheets Sheets => _sheets ??= new ExcelSheets(_workbook?.Sheets);

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

    /// <summary>
    /// 活动工作表缓存
    /// </summary>
    private IExcelWorksheet _activeWorksheet;

    /// <summary>
    /// 获取活动工作表
    /// </summary>
    public IExcelWorksheet ActiveSheet => _activeWorksheet ?? (_activeWorksheet = new ExcelWorksheet(_workbook?.ActiveSheet as MsExcel.Worksheet));

    #endregion

    #region 保护和安全
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
    public void SaveAs(string filename, int fileFormat = 0, string password = "",
                      string writeResPassword = "", bool readOnlyRecommended = false,
                      bool createBackup = false, XlSaveAsAccessMode accessMode = XlSaveAsAccessMode.xlNoChange, int conflictResolution = 2,
                      bool addToMru = true, bool local = false)
    {
        if (_workbook == null || string.IsNullOrEmpty(filename))
            return;

        try
        {
            _workbook.SaveAs(
                filename, fileFormat, password, writeResPassword, readOnlyRecommended,
                createBackup, (MsExcel.XlSaveAsAccessMode)accessMode, conflictResolution, addToMru, local
            );
        }
        catch
        {
            // 忽略保存过程中的异常
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
    /// 复制工作簿
    /// </summary>
    /// <param name="before">复制到指定工作簿之前</param>
    /// <param name="after">复制到指定工作簿之后</param>
    public void Copy(IExcelWorkbook before = null, IExcelWorkbook after = null)
    {
        // 注意：Excel中工作簿不能直接复制
        // 这里提供一个空实现以保持接口一致性
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
        catch
        {
            // 忽略发送过程中的异常
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
                try
                {
                    _worksheets[i]?.Calculate();
                }
                catch
                {
                    // 忽略单个工作表计算异常
                }
            }
        }
        catch
        {
            // 忽略计算过程中的异常
        }
    }


    /// <summary>
    /// 刷新工作簿
    /// </summary>
    public void RefreshAll()
    {
        _workbook.RefreshAll();
    }

    /// <summary>
    /// 应用自动筛选到所有工作表
    /// </summary>
    public void AutoFilterAll()
    {
        var worksheets = Sheets;
        if (worksheets != null)
        {
            for (int i = 1; i <= worksheets.Count; i++)
            {
                try
                {
                    worksheets[i]?.AutoFilter();
                }
                catch
                {
                    // 忽略单个工作表异常
                }
            }
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
            for (int i = 1; i <= worksheets.Count; i++)
            {
                try
                {
                    worksheets[i]?.ClearAll();
                }
                catch
                {
                    // 忽略单个工作表异常
                }
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
    public IEnumerable<IExcelPivotCache> PivotCaches()
    {
        var caches = _workbook?.PivotCaches();
        foreach (var item in caches)
        {
            yield return new ExcelPivotCache(item as MsExcel.PivotCache);
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
        get => (XlWindowState)(_workbook?.Windows[1]?.WindowState ?? 0);
        set
        {
            if (_workbook?.Windows[1] != null)
                _workbook.Windows[1].WindowState = (MsExcel.XlWindowState)value;
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
        get => (XlCalculation)(_workbook?.Application?.Calculation ?? 0);
        set
        {
            if (_workbook?.Application != null)
                _workbook.Application.Calculation = (MsExcel.XlCalculation)value;
        }
    }

    #endregion
}

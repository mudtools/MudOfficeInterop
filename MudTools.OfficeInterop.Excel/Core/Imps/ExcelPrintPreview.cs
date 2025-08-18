//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel PrintPreview 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.PrintPreview 相关功能的安全访问和资源管理
/// 注意：Excel没有直接的PrintPreview对象，这里通过Worksheet或Workbook的PrintPreview方法实现
/// </summary>
internal class ExcelPrintPreview : IExcelPrintPreview
{
    /// <summary>
    /// 底层的 COM 对象（可以是Worksheet、Workbook或Application）
    /// </summary>
    private object _parentObject;

    /// <summary>
    /// 父对象类型
    /// </summary>
    private enum ParentType
    {
        Worksheet,
        Workbook,
        Application
    }

    /// <summary>
    /// 父对象类型
    /// </summary>
    private ParentType _parentType;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 当前页面设置（用于预览设置）
    /// </summary>
    private MsExcel.PageSetup _pageSetup;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelPrintPreview 实例（基于Worksheet）
    /// </summary>
    /// <param name="worksheet">底层的 COM Worksheet 对象</param>
    internal ExcelPrintPreview(MsExcel.Worksheet worksheet)
    {
        _parentObject = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        _parentType = ParentType.Worksheet;
        _pageSetup = worksheet.PageSetup;
        _disposedValue = false;
    }

    /// <summary>
    /// 初始化 ExcelPrintPreview 实例（基于Workbook）
    /// </summary>
    /// <param name="workbook">底层的 COM Workbook 对象</param>
    internal ExcelPrintPreview(MsExcel.Workbook workbook)
    {
        _parentObject = workbook ?? throw new ArgumentNullException(nameof(workbook));
        _parentType = ParentType.Workbook;
        // Workbook没有直接的PageSetup，使用第一个工作表的PageSetup
        try
        {
            var firstSheet = workbook.Worksheets[1] as MsExcel.Worksheet;
            _pageSetup = firstSheet?.PageSetup;
        }
        catch
        {
            _pageSetup = null;
        }
        _disposedValue = false;
    }

    /// <summary>
    /// 初始化 ExcelPrintPreview 实例（基于Application）
    /// </summary>
    /// <param name="application">底层的 COM Application 对象</param>
    internal ExcelPrintPreview(MsExcel.Application application)
    {
        _parentObject = application ?? throw new ArgumentNullException(nameof(application));
        _parentType = ParentType.Application;
        // Application没有直接的PageSetup
        _pageSetup = null;
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
                // 释放PageSetup对象
                if (_pageSetup != null)
                    Marshal.ReleaseComObject(_pageSetup);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _parentObject = null;
            _pageSetup = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性
    /// <summary>
    /// 获取打印预览窗口的父对象
    /// </summary>
    public object Parent => _parentObject;
    #endregion

    #region 显示设置

    /// <summary>
    /// 获取或设置打印预览的缩放比例
    /// </summary>
    public int Zoom
    {
        get => (int)(_pageSetup?.Zoom ?? 100);
        set
        {
            if (_pageSetup != null)
                _pageSetup.Zoom = Math.Max(10, Math.Min(400, value));
        }
    }

    /// <summary>
    /// 获取或设置是否显示页眉
    /// </summary>
    public bool ShowHeaders
    {
        get => !string.IsNullOrEmpty(_pageSetup?.CenterHeader) ||
               !string.IsNullOrEmpty(_pageSetup?.LeftHeader) ||
               !string.IsNullOrEmpty(_pageSetup?.RightHeader);
        set
        {
            if (_pageSetup != null && value)
            {
                if (string.IsNullOrEmpty(_pageSetup.CenterHeader))
                    _pageSetup.CenterHeader = "&A";
            }
        }
    }

    /// <summary>
    /// 获取或设置是否显示页脚
    /// </summary>
    public bool ShowFooters
    {
        get => !string.IsNullOrEmpty(_pageSetup?.CenterFooter) ||
               !string.IsNullOrEmpty(_pageSetup?.LeftFooter) ||
               !string.IsNullOrEmpty(_pageSetup?.RightFooter);
        set
        {
            if (_pageSetup != null && value)
            {
                if (string.IsNullOrEmpty(_pageSetup.CenterFooter))
                    _pageSetup.CenterFooter = "第 &P 页，共 &N 页";
            }
        }
    }

    /// <summary>
    /// 获取或设置是否显示网格线
    /// </summary>
    public bool ShowGridlines
    {
        get => _pageSetup?.PrintGridlines ?? false;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintGridlines = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示行列标题
    /// </summary>
    public bool ShowHeadings
    {
        get => _pageSetup?.PrintHeadings ?? false;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintHeadings = value;
        }
    }


    /// <summary>
    /// 获取或设置是否显示注释
    /// </summary>
    public int ShowComments
    {
        get => _pageSetup != null ? Convert.ToInt32(_pageSetup.PrintComments) : 0;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintComments = (MsExcel.XlPrintLocation)value;
        }
    }

    #endregion

    #region 页面设置

    /// <summary>
    /// 获取或设置页面方向
    /// </summary>
    public int Orientation
    {
        get => _pageSetup != null ? Convert.ToInt32(_pageSetup.Orientation) : 0;
        set
        {
            if (_pageSetup != null)
                _pageSetup.Orientation = (MsExcel.XlPageOrientation)value;
        }
    }

    /// <summary>
    /// 获取或设置纸张大小
    /// </summary>
    public int PaperSize
    {
        get => _pageSetup != null ? Convert.ToInt32(_pageSetup.PaperSize) : 0;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PaperSize = (MsExcel.XlPaperSize)value;
        }
    }
    #endregion

    #region 页边距设置

    /// <summary>
    /// 获取或设置左边距（英寸）
    /// </summary>
    public double LeftMargin
    {
        get => _pageSetup?.LeftMargin ?? 0.75;
        set
        {
            if (_pageSetup != null)
                _pageSetup.LeftMargin = value;
        }
    }

    /// <summary>
    /// 获取或设置右边距（英寸）
    /// </summary>
    public double RightMargin
    {
        get => _pageSetup?.RightMargin ?? 0.75;
        set
        {
            if (_pageSetup != null)
                _pageSetup.RightMargin = value;
        }
    }

    /// <summary>
    /// 获取或设置上边距（英寸）
    /// </summary>
    public double TopMargin
    {
        get => _pageSetup?.TopMargin ?? 1.0;
        set
        {
            if (_pageSetup != null)
                _pageSetup.TopMargin = value;
        }
    }

    /// <summary>
    /// 获取或设置下边距（英寸）
    /// </summary>
    public double BottomMargin
    {
        get => _pageSetup?.BottomMargin ?? 1.0;
        set
        {
            if (_pageSetup != null)
                _pageSetup.BottomMargin = value;
        }
    }

    /// <summary>
    /// 获取或设置页眉边距（英寸）
    /// </summary>
    public double HeaderMargin
    {
        get => _pageSetup?.HeaderMargin ?? 0.5;
        set
        {
            if (_pageSetup != null)
                _pageSetup.HeaderMargin = value;
        }
    }

    /// <summary>
    /// 获取或设置页脚边距（英寸）
    /// </summary>
    public double FooterMargin
    {
        get => _pageSetup?.FooterMargin ?? 0.5;
        set
        {
            if (_pageSetup != null)
                _pageSetup.FooterMargin = value;
        }
    }

    #endregion

    #region 页眉页脚设置

    /// <summary>
    /// 获取或设置左页眉内容
    /// </summary>
    public string LeftHeader
    {
        get => _pageSetup?.LeftHeader?.ToString();
        set
        {
            if (_pageSetup != null && value != null)
                _pageSetup.LeftHeader = value;
        }
    }

    /// <summary>
    /// 获取或设置中页眉内容
    /// </summary>
    public string CenterHeader
    {
        get => _pageSetup?.CenterHeader?.ToString();
        set
        {
            if (_pageSetup != null && value != null)
                _pageSetup.CenterHeader = value;
        }
    }

    /// <summary>
    /// 获取或设置右页眉内容
    /// </summary>
    public string RightHeader
    {
        get => _pageSetup?.RightHeader?.ToString();
        set
        {
            if (_pageSetup != null && value != null)
                _pageSetup.RightHeader = value;
        }
    }

    /// <summary>
    /// 获取或设置左页脚内容
    /// </summary>
    public string LeftFooter
    {
        get => _pageSetup?.LeftFooter?.ToString();
        set
        {
            if (_pageSetup != null && value != null)
                _pageSetup.LeftFooter = value;
        }
    }

    /// <summary>
    /// 获取或设置中页脚内容
    /// </summary>
    public string CenterFooter
    {
        get => _pageSetup?.CenterFooter?.ToString();
        set
        {
            if (_pageSetup != null && value != null)
                _pageSetup.CenterFooter = value;
        }
    }

    /// <summary>
    /// 获取或设置右页脚内容
    /// </summary>
    public string RightFooter
    {
        get => _pageSetup?.RightFooter?.ToString();
        set
        {
            if (_pageSetup != null && value != null)
                _pageSetup.RightFooter = value;
        }
    }

    #endregion

    #region 操作方法

    /// <summary>
    /// 显示打印预览窗口
    /// </summary>
    /// <param name="enableChanges">是否允许在预览中进行更改</param>
    public void Show(bool enableChanges = true)
    {
        try
        {
            switch (_parentType)
            {
                case ParentType.Worksheet:
                    var worksheet = _parentObject as MsExcel.Worksheet;
                    worksheet?.PrintPreview(enableChanges);
                    break;
                case ParentType.Workbook:
                    var workbook = _parentObject as MsExcel.Workbook;
                    workbook?.PrintPreview(enableChanges);
                    break;
                case ParentType.Application:
                    var application = _parentObject as MsExcel.Application;
                    application?.ThisWorkbook?.PrintPreview(enableChanges);
                    break;
            }
        }
        catch
        {
            // 忽略显示预览过程中的异常
        }
    }

    /// <summary>
    /// 刷新打印预览显示
    /// </summary>
    public void Refresh()
    {
        // 重新显示预览以实现刷新效果
        Show();
    }

    /// <summary>
    /// 打印当前预览的内容
    /// </summary>
    /// <param name="copies">打印份数</param>
    /// <param name="collate">是否逐份打印</param>
    public void Print(int copies = 1, bool collate = true)
    {
        try
        {
            switch (_parentType)
            {
                case ParentType.Worksheet:
                    var worksheet = _parentObject as MsExcel.Worksheet;
                    worksheet?.PrintOut(
                        Type.Missing, Type.Missing, copies, collate,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing
                    );
                    break;
                case ParentType.Workbook:
                    var workbook = _parentObject as MsExcel.Workbook;
                    workbook?.PrintOut(
                        Type.Missing, Type.Missing, copies, collate,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing
                    );
                    break;
                case ParentType.Application:
                    var application = _parentObject as MsExcel.Application;
                    application?.ThisWorkbook?.PrintOut(
                        Type.Missing, Type.Missing, copies, collate,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing
                    );
                    break;
            }
        }
        catch
        {
            // 忽略打印过程中的异常
        }
    }

    /// <summary>
    /// 导出预览为PDF文件
    /// </summary>
    /// <param name="filename">PDF文件路径</param>
    public void ExportToPDF(string filename)
    {
        if (string.IsNullOrEmpty(filename)) return;

        try
        {
            switch (_parentType)
            {
                case ParentType.Worksheet:
                    var worksheet = _parentObject as MsExcel.Worksheet;
                    worksheet?.ExportAsFixedFormat(
                        MsExcel.XlFixedFormatType.xlTypePDF,
                        filename,
                        MsExcel.XlFixedFormatQuality.xlQualityStandard,
                        true, true, Type.Missing, Type.Missing, false, Type.Missing
                    );
                    break;
                case ParentType.Workbook:
                    var workbook = _parentObject as MsExcel.Workbook;
                    workbook?.ExportAsFixedFormat(
                        MsExcel.XlFixedFormatType.xlTypePDF,
                        filename,
                        MsExcel.XlFixedFormatQuality.xlQualityStandard,
                        true, true, Type.Missing, Type.Missing, false, Type.Missing
                    );
                    break;
                case ParentType.Application:
                    var application = _parentObject as MsExcel.Application;
                    application?.ThisWorkbook?.ExportAsFixedFormat(
                        MsExcel.XlFixedFormatType.xlTypePDF,
                        filename,
                        MsExcel.XlFixedFormatQuality.xlQualityStandard,
                        true, true, Type.Missing, Type.Missing, false, Type.Missing
                    );
                    break;
            }
        }
        catch
        {
            // 忽略导出过程中的异常
        }
    }

    #endregion

    #region 高级功能

    /// <summary>
    /// 获取或设置是否显示黑白预览
    /// </summary>
    public bool BlackAndWhite
    {
        get => _pageSetup?.BlackAndWhite ?? false;
        set
        {
            if (_pageSetup != null)
                _pageSetup.BlackAndWhite = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示注释预览
    /// </summary>
    public bool PrintNotes
    {
        get => _pageSetup?.PrintNotes ?? false;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintNotes = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示网格线预览
    /// </summary>
    public bool PrintGridlines
    {
        get => _pageSetup?.PrintGridlines ?? false;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintGridlines = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示行列标题预览
    /// </summary>
    public bool PrintHeadings
    {
        get => _pageSetup?.PrintHeadings ?? false;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintHeadings = value;
        }
    }

    /// <summary>
    /// 获取或设置打印区域预览
    /// </summary>
    public string PrintArea
    {
        get => _pageSetup?.PrintArea?.ToString();
        set
        {
            if (_pageSetup != null && value != null)
                _pageSetup.PrintArea = value;
        }
    }

    /// <summary>
    /// 获取或设置打印标题行预览
    /// </summary>
    public string PrintTitleRows
    {
        get => _pageSetup?.PrintTitleRows?.ToString();
        set
        {
            if (_pageSetup != null && value != null)
                _pageSetup.PrintTitleRows = value;
        }
    }

    /// <summary>
    /// 获取或设置打印标题列预览
    /// </summary>
    public string PrintTitleColumns
    {
        get => _pageSetup?.PrintTitleColumns?.ToString();
        set
        {
            if (_pageSetup != null && value != null)
                _pageSetup.PrintTitleColumns = value;
        }
    }

    #endregion
}
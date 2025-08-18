//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel PageSetup 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.PageSetup 对象的安全访问和资源管理
/// </summary>
internal class ExcelPageSetup : IExcelPageSetup
{
    /// <summary>
    /// 底层的 COM PageSetup 对象
    /// </summary>
    private MsExcel.PageSetup _pageSetup;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelPageSetup 实例
    /// </summary>
    /// <param name="pageSetup">底层的 COM PageSetup 对象</param>
    internal ExcelPageSetup(MsExcel.PageSetup pageSetup)
    {
        _pageSetup = pageSetup;
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
                // 释放底层COM对象
                if (_pageSetup != null)
                    Marshal.ReleaseComObject(_pageSetup);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _pageSetup = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 页面设置

    public IExcelPages? Pages => _pageSetup != null ? new ExcelPages(_pageSetup.Pages) : null;

    public IExcelPage? EvenPage => _pageSetup != null ? new ExcelPage(_pageSetup.EvenPage) : null;

    public IExcelPage? FirstPage => _pageSetup != null ? new ExcelPage(_pageSetup.FirstPage) : null;

    public IExcelGraphic? CenterHeaderPicture
    {
        get
        {
            return _pageSetup != null ? new ExcelGraphic(_pageSetup.CenterHeaderPicture) : null;
        }
    }


    public IExcelGraphic CenterFooterPicture
    {
        get
        {
            return _pageSetup != null ? new ExcelGraphic(_pageSetup.CenterFooterPicture) : null;
        }
    }

    public IExcelGraphic LeftHeaderPicture
    {
        get
        {
            return _pageSetup != null ? new ExcelGraphic(_pageSetup.LeftHeaderPicture) : null;
        }
    }

    public IExcelGraphic LeftFooterPicture
    {
        get
        {
            return _pageSetup != null ? new ExcelGraphic(_pageSetup.LeftFooterPicture) : null;
        }
    }

    public IExcelGraphic RightHeaderPicture
    {
        get
        {
            return _pageSetup != null ? new ExcelGraphic(_pageSetup.RightHeaderPicture) : null;
        }
    }

    public IExcelGraphic RightFooterPicture
    {
        get
        {
            return _pageSetup != null ? new ExcelGraphic(_pageSetup.RightFooterPicture) : null;
        }
    }

    /// <summary>
    /// 获取或设置页面方向（纵向或横向）
    /// </summary>
    public XlPageOrientation Orientation
    {
        get => _pageSetup != null ? (XlPageOrientation)_pageSetup.Orientation : XlPageOrientation.xlLandscape;
        set
        {
            if (_pageSetup != null)
                _pageSetup.Orientation = (MsExcel.XlPageOrientation)value;
        }
    }

    /// <summary>
    /// 获取或设置纸张大小
    /// </summary>
    public XlPaperSize PaperSize
    {
        get => _pageSetup != null ? (XlPaperSize)_pageSetup.PaperSize : XlPaperSize.xlPaperA4;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PaperSize = (MsExcel.XlPaperSize)value;
        }
    }

    /// <summary>
    /// 获取或设置页面缩放比例
    /// </summary>
    public int Zoom
    {
        get => (int)(_pageSetup != null ? _pageSetup.Zoom : 0);
        set
        {
            if (_pageSetup != null)
                _pageSetup.Zoom = value;
        }
    }

    /// <summary>
    /// 获取或设置是否适合页面宽度
    /// </summary>
    public int FitToPagesWide
    {
        get => (int)(_pageSetup != null ? _pageSetup.FitToPagesWide : 0);
        set
        {
            if (_pageSetup != null)
                _pageSetup.FitToPagesWide = value;
        }
    }

    /// <summary>
    /// 获取或设置是否适合页面高度
    /// </summary>
    public int FitToPagesTall
    {
        get => (int)(_pageSetup != null ? _pageSetup.FitToPagesTall : 0);
        set
        {
            if (_pageSetup != null)
                _pageSetup.FitToPagesTall = value;
        }
    }

    /// <summary>
    /// 获取或设置是否为黑白打印
    /// </summary>
    public bool BlackAndWhite
    {
        get => _pageSetup != null && _pageSetup.BlackAndWhite;
        set
        {
            if (_pageSetup != null)
                _pageSetup.BlackAndWhite = value;
        }
    }

    /// <summary>
    /// 获取或设置是否为单色打印
    /// </summary>
    public XlPrintLocation PrintComments
    {
        get => _pageSetup != null ? (XlPrintLocation)_pageSetup.PrintComments : XlPrintLocation.xlPrintSheetEnd;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintComments = (MsExcel.XlPrintLocation)value;
        }
    }

    /// <summary>
    /// 获取或设置打印错误处理方式
    /// </summary>
    public XlPrintErrors PrintErrors
    {
        get => _pageSetup != null ? (XlPrintErrors)_pageSetup.PrintErrors : XlPrintErrors.xlPrintErrorsDisplayed;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintErrors = (MsExcel.XlPrintErrors)value;
        }
    }

    #endregion

    #region 页边距设置

    /// <summary>
    /// 获取或设置左边距（英寸）
    /// </summary>
    public double LeftMargin
    {
        get => _pageSetup?.LeftMargin ?? 0;
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
        get => _pageSetup?.RightMargin ?? 0;
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
        get => _pageSetup?.TopMargin ?? 0;
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
        get => _pageSetup?.BottomMargin ?? 0;
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
        get => _pageSetup?.HeaderMargin ?? 0;
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
        get => _pageSetup?.FooterMargin ?? 0;
        set
        {
            if (_pageSetup != null)
                _pageSetup.FooterMargin = value;
        }
    }

    /// <summary>
    /// 获取或设置居中方式（水平居中）
    /// </summary>
    public bool CenterHorizontally
    {
        get => _pageSetup != null && _pageSetup.CenterHorizontally;
        set
        {
            if (_pageSetup != null)
                _pageSetup.CenterHorizontally = value;
        }
    }

    /// <summary>
    /// 获取或设置居中方式（垂直居中）
    /// </summary>
    public bool CenterVertically
    {
        get => _pageSetup != null && _pageSetup.CenterVertically;
        set
        {
            if (_pageSetup != null)
                _pageSetup.CenterVertically = value;
        }
    }

    #endregion

    #region 页眉页脚设置

    /// <summary>
    /// 获取或设置左页眉内容
    /// </summary>
    public string LeftHeader
    {
        get => _pageSetup?.LeftHeader;
        set
        {
            if (_pageSetup != null)
                _pageSetup.LeftHeader = value;
        }
    }

    /// <summary>
    /// 获取或设置中页眉内容
    /// </summary>
    public string CenterHeader
    {
        get => _pageSetup?.CenterHeader;
        set
        {
            if (_pageSetup != null)
                _pageSetup.CenterHeader = value;
        }
    }

    /// <summary>
    /// 获取或设置右页眉内容
    /// </summary>
    public string RightHeader
    {
        get => _pageSetup?.RightHeader;
        set
        {
            if (_pageSetup != null)
                _pageSetup.RightHeader = value;
        }
    }

    /// <summary>
    /// 获取或设置左页脚内容
    /// </summary>
    public string LeftFooter
    {
        get => _pageSetup?.LeftFooter;
        set
        {
            if (_pageSetup != null)
                _pageSetup.LeftFooter = value;
        }
    }

    /// <summary>
    /// 获取或设置中页脚内容
    /// </summary>
    public string CenterFooter
    {
        get => _pageSetup?.CenterFooter;
        set
        {
            if (_pageSetup != null)
                _pageSetup.CenterFooter = value;
        }
    }

    /// <summary>
    /// 获取或设置右页脚内容
    /// </summary>
    public string RightFooter
    {
        get => _pageSetup?.RightFooter;
        set
        {
            if (_pageSetup != null)
                _pageSetup.RightFooter = value;
        }
    }

    #endregion

    #region 打印选项
    public bool OddAndEvenPagesHeaderFooter
    {
        get => _pageSetup != null ? _pageSetup.OddAndEvenPagesHeaderFooter : false;
        set => _pageSetup.OddAndEvenPagesHeaderFooter = value;
    }

    public bool ScaleWithDocHeaderFooter
    {
        get => _pageSetup != null ? _pageSetup.ScaleWithDocHeaderFooter : false;
        set => _pageSetup.ScaleWithDocHeaderFooter = value;
    }

    public bool AlignMarginsHeaderFooter
    {
        get => _pageSetup != null ? _pageSetup.AlignMarginsHeaderFooter : false;
        set => _pageSetup.AlignMarginsHeaderFooter = value;
    }

    public bool Draft
    {
        get => _pageSetup != null ? _pageSetup.Draft : false;
        set => _pageSetup.Draft = value;
    }

    public XlOrder Order
    {
        get => _pageSetup != null ? (XlOrder)_pageSetup.Order : XlOrder.xlDownThenOver;
        set => _pageSetup.Order = (MsExcel.XlOrder)value;
    }

    public object? PrintQuality
    {
        get => _pageSetup != null ? _pageSetup.PrintQuality : null;
        set => _pageSetup.PrintQuality = value;
    }

    /// <summary>
    /// 获取或设置是否打印网格线
    /// </summary>
    public bool PrintGridlines
    {
        get => _pageSetup != null && _pageSetup.PrintGridlines;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintGridlines = value;
        }
    }

    /// <summary>
    /// 获取或设置是否打印行列标题
    /// </summary>
    public bool PrintHeadings
    {
        get => _pageSetup != null && _pageSetup.PrintHeadings;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintHeadings = value;
        }
    }

    /// <summary>
    /// 获取或设置是否打印注释
    /// </summary>
    public bool PrintNotes
    {
        get => _pageSetup != null && _pageSetup.PrintNotes;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintNotes = value;
        }
    }

    /// <summary>
    /// 获取或设置是否打印标题行
    /// </summary>
    public string PrintTitleRows
    {
        get => _pageSetup?.PrintTitleRows;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintTitleRows = value;
        }
    }

    /// <summary>
    /// 获取或设置是否打印标题列
    /// </summary>
    public string PrintTitleColumns
    {
        get => _pageSetup?.PrintTitleColumns;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintTitleColumns = value;
        }
    }

    /// <summary>
    /// 获取或设置打印区域
    /// </summary>
    public string PrintArea
    {
        get => _pageSetup?.PrintArea;
        set
        {
            if (_pageSetup != null && value != null)
                _pageSetup.PrintArea = value;
        }
    }

    /// <summary>
    /// 获取或设置是否从第一页开始编号
    /// </summary>
    public int FirstPageNumber
    {
        get => _pageSetup?.FirstPageNumber ?? 0;
        set
        {
            if (_pageSetup != null)
                _pageSetup.FirstPageNumber = value;
        }
    }

    /// <summary>
    /// 获取或设置是否不同奇偶页页眉页脚
    /// </summary>
    public bool DifferentFirstPageHeaderFooter
    {
        get => _pageSetup != null && _pageSetup.DifferentFirstPageHeaderFooter;
        set
        {
            if (_pageSetup != null)
                _pageSetup.DifferentFirstPageHeaderFooter = value;
        }
    }

    #endregion

    #region 页面编号和日期

    /// <summary>
    /// 获取或设置是否显示页码
    /// </summary>
    public bool ShowPageNumbers
    {
        get
        {
            // 通过检查页脚是否包含页码代码来判断
            return (_pageSetup?.CenterFooter?.Contains("&P") ?? false) ||
                   (_pageSetup?.LeftFooter?.Contains("&P") ?? false) ||
                   (_pageSetup?.RightFooter?.Contains("&P") ?? false);
        }
        set
        {
            if (_pageSetup != null && value)
            {
                // 如果启用，添加页码到中页脚
                if (string.IsNullOrEmpty(_pageSetup.CenterFooter))
                    _pageSetup.CenterFooter = "第 &P 页";
            }
        }
    }

    /// <summary>
    /// 获取或设置是否显示日期
    /// </summary>
    public bool ShowDate
    {
        get
        {
            // 通过检查页眉是否包含日期代码来判断
            return (_pageSetup?.CenterHeader?.Contains("&D") ?? false) ||
                   (_pageSetup?.LeftHeader?.Contains("&D") ?? false) ||
                   (_pageSetup?.RightHeader?.Contains("&D") ?? false);
        }
        set
        {
            if (_pageSetup != null && value)
            {
                // 如果启用，添加日期到右页眉
                if (string.IsNullOrEmpty(_pageSetup.RightHeader))
                    _pageSetup.RightHeader = "&D";
            }
        }
    }

    /// <summary>
    /// 获取或设置是否显示时间
    /// </summary>
    public bool ShowTime
    {
        get
        {
            // 通过检查页眉是否包含时间代码来判断
            return (_pageSetup?.CenterHeader?.Contains("&T") ?? false) ||
                   (_pageSetup?.LeftHeader?.Contains("&T") ?? false) ||
                   (_pageSetup?.RightHeader?.Contains("&T") ?? false);
        }
        set
        {
            if (_pageSetup != null && value)
            {
                // 如果启用，添加时间到右页眉
                if (string.IsNullOrEmpty(_pageSetup.RightHeader))
                    _pageSetup.RightHeader = "&T";
            }
        }
    }

    /// <summary>
    /// 获取或设置是否显示文件名
    /// </summary>
    public bool ShowFileName
    {
        get
        {
            // 通过检查页眉是否包含文件名代码来判断
            return (_pageSetup?.CenterHeader?.Contains("&F") ?? false) ||
                   (_pageSetup?.LeftHeader?.Contains("&F") ?? false) ||
                   (_pageSetup?.RightHeader?.Contains("&F") ?? false);
        }
        set
        {
            if (_pageSetup != null && value)
            {
                // 如果启用，添加文件名到左页眉
                if (string.IsNullOrEmpty(_pageSetup.LeftHeader))
                    _pageSetup.LeftHeader = "&F";
            }
        }
    }

    /// <summary>
    /// 获取或设置是否显示工作表名
    /// </summary>
    public bool ShowSheetName
    {
        get
        {
            // 通过检查页眉是否包含工作表名代码来判断
            return (_pageSetup?.CenterHeader?.Contains("&A") ?? false) ||
                   (_pageSetup?.LeftHeader?.Contains("&A") ?? false) ||
                   (_pageSetup?.RightHeader?.Contains("&A") ?? false);
        }
        set
        {
            if (_pageSetup != null && value)
            {
                // 如果启用，添加工作表名到中页眉
                if (string.IsNullOrEmpty(_pageSetup.CenterHeader))
                    _pageSetup.CenterHeader = "&A";
            }
        }
    }

    /// <summary>
    /// 获取或设置是否显示路径
    /// </summary>
    public bool ShowPath
    {
        get
        {
            // 通过检查页眉是否包含路径代码来判断
            return (_pageSetup?.CenterHeader?.Contains("&Z") ?? false) ||
                   (_pageSetup?.LeftHeader?.Contains("&Z") ?? false) ||
                   (_pageSetup?.RightHeader?.Contains("&Z") ?? false);
        }
        set
        {
            if (_pageSetup != null && value)
            {
                // 如果启用，添加路径到左页眉
                if (string.IsNullOrEmpty(_pageSetup.LeftHeader))
                    _pageSetup.LeftHeader = "&Z";
            }
        }
    }

    #endregion

    #region 操作方法

    /// <summary>
    /// 应用页面设置
    /// </summary>
    public void Apply()
    {
        // PageSetup 通常会自动应用设置
        // 这里提供一个空实现以保持接口一致性
    }

    /// <summary>
    /// 重置页面设置为默认值
    /// </summary>
    public void Reset()
    {
        if (_pageSetup == null) return;

        try
        {
            // 重置常见设置为默认值
            _pageSetup.Orientation = MsExcel.XlPageOrientation.xlPortrait;
            _pageSetup.PaperSize = MsExcel.XlPaperSize.xlPaperLetter;
            _pageSetup.Zoom = 100;
            _pageSetup.FitToPagesWide = 1;
            _pageSetup.FitToPagesTall = 1;
            _pageSetup.BlackAndWhite = false;
            _pageSetup.PrintComments = MsExcel.XlPrintLocation.xlPrintNoComments;
            _pageSetup.PrintErrors = MsExcel.XlPrintErrors.xlPrintErrorsDisplayed;

            // 重置页边距
            _pageSetup.LeftMargin = 0.75;
            _pageSetup.RightMargin = 0.75;
            _pageSetup.TopMargin = 1.0;
            _pageSetup.BottomMargin = 1.0;
            _pageSetup.HeaderMargin = 0.5;
            _pageSetup.FooterMargin = 0.5;
            _pageSetup.CenterHorizontally = false;
            _pageSetup.CenterVertically = false;

            // 清空页眉页脚
            _pageSetup.LeftHeader = "";
            _pageSetup.CenterHeader = "";
            _pageSetup.RightHeader = "";
            _pageSetup.LeftFooter = "";
            _pageSetup.CenterFooter = "";
            _pageSetup.RightFooter = "";

            // 重置打印选项
            _pageSetup.PrintGridlines = false;
            _pageSetup.PrintHeadings = false;
            _pageSetup.PrintNotes = false;
            _pageSetup.PrintTitleRows = "";
            _pageSetup.PrintTitleColumns = "";
            _pageSetup.PrintArea = "";
            _pageSetup.FirstPageNumber = 0;
            _pageSetup.DifferentFirstPageHeaderFooter = false;
        }
        catch
        {
            // 忽略重置过程中的异常
        }
    }

    /// <summary>
    /// 复制页面设置
    /// </summary>
    /// <param name="source">源页面设置对象</param>
    public void Copy(IExcelPageSetup source)
    {
        if (_pageSetup == null || source == null) return;

        try
        {
            var excelPageSetup = source as ExcelPageSetup;
            if (excelPageSetup?._pageSetup == null) return;

            // 复制所有设置
            _pageSetup.Orientation = excelPageSetup._pageSetup.Orientation;
            _pageSetup.PaperSize = excelPageSetup._pageSetup.PaperSize;
            _pageSetup.Zoom = excelPageSetup._pageSetup.Zoom;
            _pageSetup.FitToPagesWide = excelPageSetup._pageSetup.FitToPagesWide;
            _pageSetup.FitToPagesTall = excelPageSetup._pageSetup.FitToPagesTall;
            _pageSetup.BlackAndWhite = excelPageSetup._pageSetup.BlackAndWhite;
            _pageSetup.PrintComments = excelPageSetup._pageSetup.PrintComments;
            _pageSetup.PrintErrors = excelPageSetup._pageSetup.PrintErrors;

            _pageSetup.LeftMargin = excelPageSetup._pageSetup.LeftMargin;
            _pageSetup.RightMargin = excelPageSetup._pageSetup.RightMargin;
            _pageSetup.TopMargin = excelPageSetup._pageSetup.TopMargin;
            _pageSetup.BottomMargin = excelPageSetup._pageSetup.BottomMargin;
            _pageSetup.HeaderMargin = excelPageSetup._pageSetup.HeaderMargin;
            _pageSetup.FooterMargin = excelPageSetup._pageSetup.FooterMargin;
            _pageSetup.CenterHorizontally = excelPageSetup._pageSetup.CenterHorizontally;
            _pageSetup.CenterVertically = excelPageSetup._pageSetup.CenterVertically;

            _pageSetup.LeftHeader = excelPageSetup._pageSetup.LeftHeader;
            _pageSetup.CenterHeader = excelPageSetup._pageSetup.CenterHeader;
            _pageSetup.RightHeader = excelPageSetup._pageSetup.RightHeader;
            _pageSetup.LeftFooter = excelPageSetup._pageSetup.LeftFooter;
            _pageSetup.CenterFooter = excelPageSetup._pageSetup.CenterFooter;
            _pageSetup.RightFooter = excelPageSetup._pageSetup.RightFooter;

            _pageSetup.PrintGridlines = excelPageSetup._pageSetup.PrintGridlines;
            _pageSetup.PrintHeadings = excelPageSetup._pageSetup.PrintHeadings;
            _pageSetup.PrintNotes = excelPageSetup._pageSetup.PrintNotes;
            _pageSetup.PrintTitleRows = excelPageSetup._pageSetup.PrintTitleRows;
            _pageSetup.PrintTitleColumns = excelPageSetup._pageSetup.PrintTitleColumns;
            _pageSetup.PrintArea = excelPageSetup._pageSetup.PrintArea;
            _pageSetup.FirstPageNumber = excelPageSetup._pageSetup.FirstPageNumber;
            _pageSetup.DifferentFirstPageHeaderFooter = excelPageSetup._pageSetup.DifferentFirstPageHeaderFooter;
        }
        catch
        {
            // 忽略复制过程中的异常
        }
    }

    /// <summary>
    /// 获取标准页眉页脚代码
    /// </summary>
    /// <param name="type">页眉页脚类型</param>
    /// <returns>标准代码</returns>
    public string GetStandardHeaderFooterCode(int type)
    {
        return type switch
        {
            1 => "&P",// 页码
            2 => "&N",// 总页数
            3 => "&D",// 日期
            4 => "&T",// 时间
            5 => "&F",// 文件名
            6 => "&A",// 工作表名
            7 => "&Z",// 路径
            8 => "&G",// 图片
            9 => "&B",// 粗体
            10 => "&I",// 斜体
            11 => "&U",// 下划线
            12 => "&\"Arial\"",// 字体
            13 => "&12",// 字号
            _ => "",
        };
    }

    /// <summary>
    /// 设置自定义页眉页脚
    /// </summary>
    /// <param name="section">区域（1=左, 2=中, 3=右）</param>
    /// <param name="position">位置（1=页眉, 2=页脚）</param>
    /// <param name="text">文本内容</param>
    public void SetCustomHeaderFooter(int section, int position, string text)
    {
        if (_pageSetup == null || string.IsNullOrEmpty(text)) return;

        try
        {
            switch (position)
            {
                case 1: // 页眉
                    switch (section)
                    {
                        case 1: _pageSetup.LeftHeader = text; break;
                        case 2: _pageSetup.CenterHeader = text; break;
                        case 3: _pageSetup.RightHeader = text; break;
                    }
                    break;
                case 2: // 页脚
                    switch (section)
                    {
                        case 1: _pageSetup.LeftFooter = text; break;
                        case 2: _pageSetup.CenterFooter = text; break;
                        case 3: _pageSetup.RightFooter = text; break;
                    }
                    break;
            }
        }
        catch
        {
            // 忽略设置过程中的异常
        }
    }

    #endregion
}
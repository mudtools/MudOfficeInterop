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
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 当前页面设置（用于预览设置）
    /// </summary>
    private MsExcel.PageSetup? _pageSetup;

    #region 构造函数和释放
    /// <summary>
    /// 初始化 ExcelPrintPreview 实例（基于Workbook）
    /// </summary>
    /// <param name="pageSetup">底层的 COM PageSetup 对象</param>
    internal ExcelPrintPreview(MsExcel.PageSetup pageSetup)
    {
        _pageSetup = pageSetup ?? throw new ArgumentNullException(nameof(pageSetup));
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
            // 释放PageSetup对象
            if (_pageSetup != null)
                Marshal.ReleaseComObject(_pageSetup);
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
    public object? Parent => _pageSetup.Parent;
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
    public XlPrintLocation ShowComments
    {
        get => _pageSetup != null ? _pageSetup.PrintComments.EnumConvert(XlPrintLocation.xlPrintNoComments) : XlPrintLocation.xlPrintNoComments;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PrintComments = value.EnumConvert(MsExcel.XlPrintLocation.xlPrintNoComments);
        }
    }

    #endregion

    #region 页面设置  

    /// <summary>
    /// 获取或设置纸张大小
    /// </summary>
    public XlPaperSize PaperSize
    {
        get => _pageSetup != null ? _pageSetup.PaperSize.EnumConvert(XlPaperSize.xlPaperA4) : XlPaperSize.xlPaperA4;
        set
        {
            if (_pageSetup != null)
                _pageSetup.PaperSize = value.EnumConvert(MsExcel.XlPaperSize.xlPaperA4);
        }
    }

    /// <summary>
    /// 获取或设置页面方向
    /// </summary>
    public XlPageOrientation Orientation
    {
        get => _pageSetup != null ? _pageSetup.Orientation.EnumConvert(XlPageOrientation.xlPortrait) : XlPageOrientation.xlPortrait;
        set
        {
            if (_pageSetup != null)
                _pageSetup.Orientation = value.EnumConvert(MsExcel.XlPageOrientation.xlPortrait);
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
        get => _pageSetup?.LeftHeader;
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
        get => _pageSetup?.CenterHeader;
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
        get => _pageSetup?.RightHeader;
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
        get => _pageSetup?.LeftFooter;
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
        get => _pageSetup?.CenterFooter;
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
        get => _pageSetup?.RightFooter;
        set
        {
            if (_pageSetup != null && value != null)
                _pageSetup.RightFooter = value;
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
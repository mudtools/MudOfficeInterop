//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel PlotArea 对象的二次封装实现类
/// 实现 IExcelPlotArea 接口
/// </summary>
internal class ExcelPlotArea : IExcelPlotArea
{
    private MsExcel.PlotArea? _plotArea;
    private bool _disposedValue = false;

    internal ExcelPlotArea(MsExcel.PlotArea plotArea)
    {
        _plotArea = plotArea ?? throw new ArgumentNullException(nameof(plotArea));
    }

    #region 基础属性
    public string Name
    {
        get => _plotArea.Name;
    }

    public object Parent => _plotArea.Parent;

    public IExcelApplication Application => new ExcelApplication(_plotArea.Application);
    #endregion

    #region 位置和大小
    public double Left
    {
        get => (double)_plotArea.Left;
        set => _plotArea.Left = value;
    }

    public double Top
    {
        get => (double)_plotArea.Top;
        set => _plotArea.Top = value;
    }

    public double Width
    {
        get => (double)_plotArea.Width;
        set => _plotArea.Width = value;
    }

    public double Height
    {
        get => (double)_plotArea.Height;
        set => _plotArea.Height = value;
    }

    public double InsideLeft
    {
        get => _plotArea.InsideLeft;
        set => _plotArea.InsideLeft = value;
    }

    public double InsideTop
    {
        get => _plotArea.InsideTop;
        set => _plotArea.InsideTop = value;
    }

    public double InsideWidth
    {
        get => _plotArea.InsideWidth;
        set => _plotArea.InsideWidth = value;
    }

    public double InsideHeight
    {
        get => _plotArea.InsideHeight;
        set => _plotArea.InsideHeight = value;
    }
    #endregion

    #region 格式设置
    public IExcelChartFormat Format => new ExcelChartFormat(_plotArea.Format);
    public IExcelFillFormat Fill => new ExcelFillFormat(_plotArea.Format.Fill);
    public IExcelBorder Border => new ExcelBorder(_plotArea.Border);
    #endregion

    #region 操作方法 

    public void ClearFormats()
    {
        _plotArea.ClearFormats();
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放底层COM对象
                if (_plotArea != null)
                    Marshal.ReleaseComObject(_plotArea);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _plotArea = null;
        }
        _disposedValue = true;
    }

    ~ExcelPlotArea()
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

//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Legend 对象的二次封装实现类
/// 实现 IExcelLegend 接口
/// </summary>
internal class ExcelLegend : IExcelLegend
{
    private MsExcel.Legend _legend;
    private bool _disposedValue = false;

    internal ExcelLegend(MsExcel.Legend legend)
    {
        _legend = legend ?? throw new ArgumentNullException(nameof(legend));
    }

    #region 基础属性
    public string Name
    {
        get => _legend.Name;
    }


    public object Parent => _legend.Parent;

    public IExcelApplication Application => new ExcelApplication(_legend.Application);
    #endregion

    #region 位置和大小
    public double Left
    {
        get => (double)_legend.Left;
        set => _legend.Left = value;
    }

    public double Top
    {
        get => (double)_legend.Top;
        set => _legend.Top = value;
    }

    public double Width
    {
        get => (double)_legend.Width;
        set => _legend.Width = value;
    }

    public double Height
    {
        get => (double)_legend.Height;
        set => _legend.Height = value;
    }
    #endregion

    #region 格式设置
    public IExcelFont Font => new ExcelFont(_legend.Font);

    public bool AutoScaleFont
    {
        get => Convert.ToBoolean(_legend.AutoScaleFont);
        set => _legend.AutoScaleFont = value;
    }

    public IExcelChartFillFormat Fill => new ExcelChartFillFormat(_legend.Fill);

    public IExcelBorder Border => new ExcelBorder(_legend.Border);

    public XlLegendPosition Position
    {
        get => (XlLegendPosition)_legend.Position;
        set => _legend.Position = (MsExcel.XlLegendPosition)value;
    }

    /// <summary>
    /// 获取样式的内部格式对象
    /// </summary>
    public IExcelInterior Interior => new ExcelInterior(_legend.Interior);

    #endregion

    #region 操作方法
    public void Select()
    {
        _legend.Select();
    }

    public void Delete()
    {
        _legend.Delete();
    }

    public void Clear()
    {
        _legend.Clear();
    }

    public BoundingBox GetBoundingBox()
    {
        return new BoundingBox
        {
            Left = Left,
            Top = Top,
            Width = Width,
            Height = Height
        };
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
                if (_legend != null)
                    Marshal.ReleaseComObject(_legend);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _legend = null;
        }
        _disposedValue = true;
    }

    ~ExcelLegend()
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

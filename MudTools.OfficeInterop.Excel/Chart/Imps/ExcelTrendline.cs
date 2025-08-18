//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Trendline 对象的二次封装实现类
/// 实现 IExcelTrendline 接口
/// </summary>
internal class ExcelTrendline : IExcelTrendline
{
    internal MsExcel.Trendline _trendline;
    private bool _disposedValue = false;

    internal ExcelTrendline(MsExcel.Trendline trendline)
    {
        _trendline = trendline ?? throw new ArgumentNullException(nameof(trendline));
    }

    #region 基础属性
    public string Name
    {
        get => _trendline.Name;
        set => _trendline.Name = value;
    }

    public int Index => _trendline.Index;

    public object Parent => _trendline.Parent;

    public IExcelApplication Application => new ExcelApplication(_trendline.Application);

    public int Type
    {
        get => (int)_trendline.Type;
        set => _trendline.Type = (MsExcel.XlTrendlineType)value;
    }

    public int Order
    {
        get => _trendline.Order;
        set => _trendline.Order = value;
    }

    public int Period
    {
        get => _trendline.Period;
        set => _trendline.Period = value;
    }

    public int Forward
    {
        get => _trendline.Forward;
        set => _trendline.Forward = value;
    }

    public int Backward
    {
        get => _trendline.Backward;
        set => _trendline.Backward = value;
    }

    public double Intercept
    {
        get => _trendline.Intercept;
        set => _trendline.Intercept = value;
    }

    public bool DisplayEquation
    {
        get => _trendline.DisplayEquation;
        set => _trendline.DisplayEquation = value;
    }

    public bool DisplayRSquared
    {
        get => _trendline.DisplayRSquared;
        set => _trendline.DisplayRSquared = value;
    }
    #endregion

    #region 格式设置

    public IExcelBorder Border => new ExcelBorder(_trendline.Border);

    public IExcelChartFormat Format => new ExcelChartFormat(_trendline.Format);
    #endregion



    #region 操作方法
    public void Select()
    {
        _trendline.Select();
    }

    public void Delete()
    {
        _trendline.Delete();
    }

    public void ClearFormats()
    {
        _trendline.ClearFormats();
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
                if (_trendline != null)
                    Marshal.ReleaseComObject(_trendline);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _trendline = null;
        }
        _disposedValue = true;
    }

    ~ExcelTrendline()
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

//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel TickLabels 对象的二次封装实现类
/// 实现 IExcelTickLabels 接口
/// </summary>
internal class ExcelTickLabels : IExcelTickLabels
{
    private MsExcel.TickLabels _tickLabels;
    private bool _disposedValue = false;

    internal ExcelTickLabels(MsExcel.TickLabels tickLabels)
    {
        _tickLabels = tickLabels ?? throw new ArgumentNullException(nameof(tickLabels));
    }

    #region 基础属性
    public string Name => _tickLabels.Name;

    public object Parent => _tickLabels.Parent;

    public IExcelApplication Application => new ExcelApplication(_tickLabels.Application);
    #endregion

    #region 格式设置
    public IExcelFont Font => new ExcelFont(_tickLabels.Font);


    public IExcelChartFormat Format => new ExcelChartFormat(_tickLabels.Format);

    public bool AutoScaleFont
    {
        get => Convert.ToBoolean(_tickLabels.AutoScaleFont);
        set => _tickLabels.AutoScaleFont = value;
    }

    public string NumberFormat
    {
        get => _tickLabels.NumberFormat;
        set => _tickLabels.NumberFormat = value;
    }

    public bool NumberFormatLinked
    {
        get => _tickLabels.NumberFormatLinked;
        set => _tickLabels.NumberFormatLinked = value;
    }

    public string? NumberFormatLocal
    {
        get => _tickLabels.NumberFormatLocal?.ToString();
        set => _tickLabels.NumberFormatLocal = value;
    }

    public int Orientation
    {
        get => (int)_tickLabels.Orientation;
        set => _tickLabels.Orientation = (MsExcel.XlTickLabelOrientation)value;
    }

    public int ReadingOrder
    {
        get => _tickLabels.ReadingOrder;
        set => _tickLabels.ReadingOrder = value;
    }

    public int Offset
    {
        get => _tickLabels.Offset;
        set => _tickLabels.Offset = value;
    }

    public bool MultiLevel
    {
        get => _tickLabels.MultiLevel;
        set => _tickLabels.MultiLevel = value;
    }
    #endregion

    #region 操作方法
    public void Select()
    {
        _tickLabels.Select();
    }

    public void Delete()
    {
        _tickLabels.Delete();
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
                if (_tickLabels != null)
                    Marshal.ReleaseComObject(_tickLabels);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _tickLabels = null;
        }
    }

    ~ExcelTickLabels()
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

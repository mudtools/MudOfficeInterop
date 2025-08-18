//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Fill 对象的二次封装实现类
/// 实现 IExcelFill 接口
/// </summary>
internal class ExcelChartFillFormat : IExcelChartFillFormat
{
    #region 私有字段
    /// <summary>
    /// 强类型引用，用于简化访问
    /// </summary>
    private MsExcel.ChartFillFormat _chartFillFormat;

    /// <summary>
    /// 标记对象是否已被释放，用于防止重复释放
    /// </summary>
    private bool _disposedValue = false;

    #endregion

    #region 构造函数
    /// <summary>
    /// 初始化 ExcelFill 实例 (用于 ChartArea.Format.Fill, PlotArea.Format.Fill 等)
    /// </summary>
    /// <param name="chartFillFormat">要封装的 Microsoft.Office.Interop.Excel.ChartFillFormat 对象</param>
    internal ExcelChartFillFormat(MsExcel.ChartFillFormat chartFillFormat)
    {
        _chartFillFormat = chartFillFormat ?? throw new ArgumentNullException(nameof(chartFillFormat));
        _chartFillFormat = chartFillFormat;
    }

    #endregion

    #region 基础属性 (IExcelFill)

    /// <summary>
    /// 获取填充所在的父对象
    /// </summary>
    public object Parent
    {
        get
        {
            if (_chartFillFormat != null) return _chartFillFormat.Parent;
            return null;
        }
    }

    /// <summary>
    /// 获取填充对象所在的 Application 对象
    /// </summary>
    public IExcelApplication Application
    {
        get
        {
            var parent = Parent;
            if (parent is MsExcel.Chart chart)
            {
                return new ExcelApplication(chart.Application);
            }
            return null;
        }
    }

    #endregion

    #region 填充属性 (IExcelFill)

    /// <summary>
    /// 获取或设置填充的前景色 (RGB 颜色值)
    /// </summary>
    public int ForeColor => _chartFillFormat.ForeColor.RGB;

    /// <summary>
    /// 获取或设置填充的背景色 (RGB 颜色值)
    /// </summary>
    public int BackColor => _chartFillFormat.BackColor.RGB;

    /// <summary>
    /// 获取或设置填充类型
    /// </summary>
    public MsoFillType? FillType => (MsoFillType)_chartFillFormat?.Type;

    /// <summary>
    /// 获取或设置图案类型
    /// </summary>
    public MsoPatternType? Pattern => (MsoPatternType)_chartFillFormat?.Pattern;

    /// <summary>
    /// 获取或设置渐变填充的样式
    /// </summary>
    public MsoGradientStyle? GradientStyle => (MsoGradientStyle)_chartFillFormat?.GradientStyle;
    /// <summary>
    /// 获取或设置渐变填充的变体
    /// </summary>
    public int GradientVariant => _chartFillFormat.GradientVariant;
    /// <summary>
    /// 获取或设置渐变填充的颜色类型
    /// </summary>
    public MsoGradientColorType? GradientColorType => (MsoGradientColorType)_chartFillFormat?.GradientColorType;

    #endregion

    #region 操作方法 (IExcelFill)   

    /// <summary>
    /// 将填充设置为纯色
    /// </summary>
    public void SetSolid()
    {
        try
        {
            _chartFillFormat?.Solid();
        }
        catch { }
    }

    /// <summary>
    /// 将填充设置为无填充
    /// </summary>
    public void SetNoFill()
    {
        try
        {
            _chartFillFormat.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
        }
        catch { }
    }

    #endregion

    #region IDisposable Support

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                if (_chartFillFormat != null)
                    Marshal.ReleaseComObject(_chartFillFormat);
            }
            catch { }
            _chartFillFormat = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 终结器 (析构函数)
    /// </summary>
    ~ExcelChartFillFormat()
    {
        Dispose(disposing: false);
    }

    /// <summary>
    /// 公开的 Dispose 方法，实现 IDisposable 接口
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
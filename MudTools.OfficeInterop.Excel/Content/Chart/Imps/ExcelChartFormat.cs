//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel ChartFormat 对象的二次封装实现类
/// 实现 IExcelChartFormat 接口
/// </summary>
internal class ExcelChartFormat : IExcelChartFormat
{
    #region 私有字段

    /// <summary>
    /// 内部持有的 Microsoft.Office.Interop.Excel.ChartFormat 对象引用
    /// </summary>
    private MsExcel.ChartFormat? _chartFormat;

    /// <summary>
    /// 标记对象是否已被释放，用于防止重复释放
    /// </summary>
    private bool _disposedValue = false;

    #endregion

    #region 构造函数

    /// <summary>
    /// 初始化 ExcelChartFormat 实例
    /// </summary>
    /// <param name="chartFormat">要封装的 Microsoft.Office.Interop.Excel.ChartFormat 对象</param>
    /// <exception cref="ArgumentNullException">当 chartFormat 为 null 时抛出</exception>
    internal ExcelChartFormat(MsExcel.ChartFormat chartFormat)
    {
        _chartFormat = chartFormat ?? throw new ArgumentNullException(nameof(chartFormat));
    }

    #endregion

    #region 基础属性 (IExcelChartFormat)

    /// <summary>
    /// 获取 ChartFormat 对象的父对象
    /// </summary>
    public object Parent => _chartFormat.Parent;

    /// <summary>
    /// 获取 ChartFormat 对象所在的 Application 对象
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_chartFormat.Application);

    #endregion

    #region 格式设置 (IExcelChartFormat)

    /// <summary>
    /// 获取图表元素的填充格式对象
    /// </summary>
    public IExcelFillFormat Fill => new ExcelFillFormat(_chartFormat.Fill);

    /// <summary>
    /// 获取图表元素的边框线条格式对象
    /// </summary>
    public IExcelLine Line => new ExcelLine(_chartFormat.Line);

    #endregion

    #region IDisposable Support

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
                if (_chartFormat != null)
                    Marshal.ReleaseComObject(_chartFormat);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _chartFormat = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 终结器 (析构函数)，防止资源未被释放
    /// </summary>
    ~ExcelChartFormat()
    {
        Dispose(disposing: false);
    }

    /// <summary>
    /// 公开的 Dispose 方法，实现 IDisposable 接口
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        // 抑制终结器
        GC.SuppressFinalize(this);
    }

    #endregion
}
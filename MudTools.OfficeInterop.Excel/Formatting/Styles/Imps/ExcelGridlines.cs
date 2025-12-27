//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Gridlines 对象的二次封装实现类
/// 实现 IExcelGridlines 接口
/// </summary>
internal class ExcelGridlines : IExcelGridlines
{
    private MsExcel.Gridlines _gridlines;
    private bool _disposedValue = false;

    internal ExcelGridlines(MsExcel.Gridlines gridlines)
    {
        _gridlines = gridlines ?? throw new ArgumentNullException(nameof(gridlines));
    }

    #region 基础属性
    public string Name => _gridlines.Name;


    public object? Parent => _gridlines.Parent;

    public IExcelApplication? Application => new ExcelApplication(_gridlines.Application);
    #endregion

    #region 格式设置
    public IExcelBorder Border => new ExcelBorder(_gridlines.Border);

    public IExcelChartFormat Format => new ExcelChartFormat(_gridlines.Format);
    #endregion  

    #region 操作方法
    public void Select()
    {
        _gridlines.Select();
    }

    public void Delete()
    {
        try
        {
            _gridlines.Delete();
        }
        catch
        {
        }
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
                if (_gridlines != null)
                    Marshal.ReleaseComObject(_gridlines);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _gridlines = null;
        }
        _disposedValue = true;
    }

    ~ExcelGridlines()
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

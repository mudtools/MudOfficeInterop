//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel CellFormat 对象的二次封装实现类
/// 实现 IExcelCellFormat 接口
/// </summary>
internal class ExcelCellFormat : IExcelCellFormat
{
    private MsExcel.CellFormat _cellFormat;
    private bool _disposedValue = false;

    internal ExcelCellFormat(MsExcel.CellFormat cellFormat)
    {
        _cellFormat = cellFormat ?? throw new ArgumentNullException(nameof(cellFormat));
    }

    #region 基础属性
    public object Parent => _cellFormat.Parent;

    public IExcelApplication Application => new ExcelApplication(_cellFormat.Application);

    public object NumberFormat
    {
        get => _cellFormat.NumberFormat;
        set => _cellFormat.NumberFormat = value;
    }

    public XlHAlign HorizontalAlignment
    {
        get => (XlHAlign)_cellFormat.HorizontalAlignment;
        set => _cellFormat.HorizontalAlignment = (MsExcel.XlHAlign)value;
    }

    public XlVAlign VerticalAlignment
    {
        get => (XlVAlign)_cellFormat.VerticalAlignment;
        set => _cellFormat.VerticalAlignment = (MsExcel.XlVAlign)value;
    }

    public int IndentLevel
    {
        get => Convert.ToInt32(_cellFormat.IndentLevel);
        set => _cellFormat.IndentLevel = value;
    }

    public XlOrientation Orientation
    {
        get => (XlOrientation)_cellFormat.Orientation;
        set => _cellFormat.Orientation = (MsExcel.XlOrientation)value; // value 应为 MsExcel.XlOrientation 枚举值对应的 int
    }

    public bool ShrinkToFit
    {
        get => Convert.ToBoolean(_cellFormat.ShrinkToFit);
        set => _cellFormat.ShrinkToFit = value;
    }

    public bool WrapText
    {
        get => Convert.ToBoolean(_cellFormat.WrapText);
        set => _cellFormat.WrapText = value;
    }

    public bool MergeCells
    {
        get => Convert.ToBoolean(_cellFormat.MergeCells);
        set => _cellFormat.MergeCells = value;
    }

    public bool Locked
    {
        get => Convert.ToBoolean(_cellFormat.Locked);
        set => _cellFormat.Locked = value;
    }

    public bool FormulaHidden
    {
        get => Convert.ToBoolean(_cellFormat.FormulaHidden);
        set => _cellFormat.FormulaHidden = value;
    }

    public string NumberFormatLocal
    {
        get => _cellFormat.NumberFormatLocal?.ToString();
        set => _cellFormat.FormulaHidden = value;
    }
    #endregion

    #region 格式设置 (子对象)
    public IExcelFont Font => _cellFormat.Font != null ? new ExcelFont(_cellFormat.Font) : null;

    public IExcelInterior Interior => _cellFormat.Interior != null ? new ExcelInterior(_cellFormat.Interior) : null;

    public IExcelBorders Borders => _cellFormat.Borders != null ? new ExcelBorders(_cellFormat.Borders) : null;


    #endregion

    #region 操作方法
    public void Clear()
    {
        _cellFormat.Clear();
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
                if (_cellFormat != null)
                    Marshal.ReleaseComObject(_cellFormat);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _cellFormat = null;
        }

        _disposedValue = true;
    }

    ~ExcelCellFormat()
    {
        // 不要更改此代码。将清理代码放入“Dispose(bool disposing)”方法中
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        // 不要更改此代码。将清理代码放入“Dispose(bool disposing)”方法中
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}

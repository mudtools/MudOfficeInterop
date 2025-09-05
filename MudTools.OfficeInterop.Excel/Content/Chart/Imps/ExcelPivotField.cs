//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel PivotField 对象的二次封装实现类
/// 实现 IExcelPivotField 接口
/// </summary>
internal class ExcelPivotField : IExcelPivotField
{
    internal MsExcel.PivotField _pivotField; // internal for ExcelPivotFields.Delete(IExcelPivotField)
    private bool _disposedValue = false;

    internal ExcelPivotField(MsExcel.PivotField pivotField)
    {
        _pivotField = pivotField ?? throw new ArgumentNullException(nameof(pivotField));
    }

    #region 基础属性
    public string Name
    {
        get => _pivotField.Name;
        set => _pivotField.Name = value;
    }


    public object Parent => _pivotField.Parent;

    public IExcelApplication Application => new ExcelApplication(_pivotField.Application);

    public XlPivotFieldOrientation Orientation
    {
        get => (XlPivotFieldOrientation)_pivotField.Orientation;
        set => _pivotField.Orientation = (MsExcel.XlPivotFieldOrientation)value;
    }

    public int Position
    {
        get => Convert.ToInt32(_pivotField.Position);
        set => _pivotField.Position = value;
    }

    public string NumberFormat
    {
        get => _pivotField.NumberFormat;
        set => _pivotField.NumberFormat = value;
    }

    public object SourceName => _pivotField.SourceName;

    public XlConsolidationFunction Function
    {
        get => (XlConsolidationFunction)_pivotField.Function;
        set => _pivotField.Function = (MsExcel.XlConsolidationFunction)value;
    }

    public string Formula
    {
        get => _pivotField.Formula;
        set => _pivotField.Formula = value;
    }
    #endregion   

    #region 状态属性 

    public bool EnableItemSelection => _pivotField.EnableItemSelection;

    public bool IsCalculated => _pivotField.IsCalculated;

    public bool IsMemberProperty => _pivotField.IsMemberProperty;
    #endregion

    #region 图表元素 (子对象)
    public IExcelPivotItems PivotItems => new ExcelPivotItems((MsExcel.PivotItems)_pivotField.PivotItems());

    public IExcelRange DataRange => new ExcelRange(_pivotField.DataRange);

    public IExcelRange LabelRange => new ExcelRange(_pivotField.LabelRange);
    #endregion

    #region 操作方法
    public void Select(bool replace = true)
    {
        _pivotField.LabelRange.Select();
    }

    public void Delete()
    {
        _pivotField.Orientation = MsExcel.XlPivotFieldOrientation.xlHidden;
    }

    public void Clear()
    {
        _pivotField.Orientation = MsExcel.XlPivotFieldOrientation.xlHidden;
    }


    public void Copy()
    {
        _pivotField.LabelRange.Copy();
    }

    public void Cut()
    {
        _pivotField.LabelRange.Cut();
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
                if (_pivotField != null)
                    Marshal.ReleaseComObject(_pivotField);
            }
            catch
            {
            }
            _pivotField = null;
        }
        _disposedValue = true;
    }

    ~ExcelPivotField()
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

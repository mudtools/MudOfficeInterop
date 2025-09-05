//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel PivotItem 对象的二次封装实现类
/// 实现 IExcelPivotItem 接口
/// </summary>
internal class ExcelPivotItem : IExcelPivotItem
{
    internal MsExcel.PivotItem _pivotItem; // internal for ExcelPivotItems.Hide(IExcelPivotItem)
    private bool _disposedValue = false;

    internal ExcelPivotItem(MsExcel.PivotItem pivotItem)
    {
        _pivotItem = pivotItem ?? throw new ArgumentNullException(nameof(pivotItem));
    }

    #region 基础属性
    public string Name => _pivotItem.Name;

    public object Parent => _pivotItem.Parent;

    public IExcelApplication Application => new ExcelApplication(_pivotItem.Application);

    public bool Visible
    {
        get => _pivotItem.Visible;
        set => _pivotItem.Visible = value;
    }

    public string Formula
    {
        get => _pivotItem.Formula;
        set => _pivotItem.Formula = value;
    }

    public string SourceName => _pivotItem.SourceName.ToString();
    #endregion

    #region 图表元素 (子对象)
    public IExcelRange DataRange => new ExcelRange(_pivotItem.DataRange);

    public IExcelRange LabelRange => new ExcelRange(_pivotItem.LabelRange);
    #endregion

    #region 操作方法
    public void Select(bool replace = true)
    {
        _pivotItem.LabelRange.Select();
    }

    public void Hide()
    {
        Visible = false;
    }

    public void Show()
    {
        Visible = true;
    }


    public void Copy()
    {
        _pivotItem.LabelRange.Copy();
    }

    public void Cut()
    {
        _pivotItem.LabelRange.Cut();
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
                if (_pivotItem != null)
                    Marshal.ReleaseComObject(_pivotItem);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _pivotItem = null;
        }
        _disposedValue = true;
    }

    ~ExcelPivotItem()
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

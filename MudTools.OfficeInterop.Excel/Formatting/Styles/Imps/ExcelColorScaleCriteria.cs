//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel ColorScaleCriteria 集合对象的二次封装实现类
/// 实现 IExcelColorScaleCriteria 接口
/// </summary>
internal class ExcelColorScaleCriteria : IExcelColorScaleCriteria
{
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelBorders));
    private MsExcel.ColorScaleCriteria? _colorScaleCriteria;
    private bool _disposedValue = false;
    private DisposableList _disposables = [];
    internal ExcelColorScaleCriteria(MsExcel.ColorScaleCriteria colorScaleCriteria)
    {
        _colorScaleCriteria = colorScaleCriteria ?? throw new ArgumentNullException(nameof(colorScaleCriteria));
    }

    #region 基础属性
    public int Count => _colorScaleCriteria != null ? _colorScaleCriteria.Count : 0;

    public IExcelColorScaleCriterion? this[int index]
    {
        get
        {
            if (_colorScaleCriteria == null || index < 1 || index > Count)
                return null;

            try
            {
                var action = _colorScaleCriteria[index];
                var a = action != null ? new ExcelColorScaleCriterion(action) : null;
                if (a != null)
                    _disposables.Add(a);
                return a;
            }
            catch (COMException ex)
            {
                log.Error("获取指定索引的颜色刻度条件对象失败：" + ex.Message, ex);
                return null;
            }
            catch (Exception ex)
            {
                log.Error("获取指定索引的颜色刻度条件对象失败：" + ex.Message, ex);
                return null;
            }
        }
    }
    #endregion

    #region IEnumerable<IExcelColorScaleCriterion> Support
    public IEnumerator<IExcelColorScaleCriterion> GetEnumerator()
    {
        if (_colorScaleCriteria == null)
            yield break;

        for (int i = 1; i <= _colorScaleCriteria.Count; i++)
        {
            yield return new ExcelColorScaleCriterion(_colorScaleCriteria[i]);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _disposables?.Dispose();
            // 释放形状对象
            if (_colorScaleCriteria != null)
                Marshal.ReleaseComObject(_colorScaleCriteria);
            _colorScaleCriteria = null;
        }

        _disposedValue = true;
    }

    ~ExcelColorScaleCriteria()
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

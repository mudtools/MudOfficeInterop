//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// Office FileDialogFilters 集合对象的二次封装实现类
/// 实现 IOfficeFileDialogFilters 接口
/// </summary>
internal class OfficeFileDialogFilters : IOfficeFileDialogFilters
{
    private MsCore.FileDialogFilters _filters;
    private bool _disposedValue = false;

    internal OfficeFileDialogFilters(MsCore.FileDialogFilters filters)
    {
        _filters = filters ?? throw new ArgumentNullException(nameof(filters));
    }

    #region 基础属性
    public int Count => _filters.Count;

    public IOfficeFileDialogFilter this[int index]
    {
        get
        {
            if (_filters == null || index < 1 || index > Count)
                return null;

            try
            {
                var name = _filters.Item(index);
                return name != null ? new OfficeFileDialogFilter(name) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    public object Parent => _filters.Parent;

    #endregion

    #region 创建和添加
    public IOfficeFileDialogFilter Add(string description, string extensions, int position = -1)
    {
        object comPosition = position > 0 ? (object)position : System.Type.Missing;
        var newFilter = _filters.Add(description, extensions, comPosition);
        return new OfficeFileDialogFilter(newFilter);
    }
    #endregion

    #region 操作方法
    public void Clear()
    {
        _filters.Clear();
    }

    public void Delete(IOfficeFileDialogFilter filter)
    {
        if (filter is OfficeFileDialogFilter officeFilter)
        {
            try
            {
                _filters.Delete(officeFilter._filter);
            }
            catch
            {
                // Handle error
            }
        }
    }
    #endregion


    #region IEnumerable<IOfficeFileDialogFilter> Support
    public IEnumerator<IOfficeFileDialogFilter> GetEnumerator()
    {
        for (int i = 1; i <= _filters.Count; i++)
        {
            yield return new OfficeFileDialogFilter(_filters.Item(i));
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
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
            try
            {
                // 释放底层COM对象
                if (_filters != null)
                    Marshal.ReleaseComObject(_filters);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _filters = null;
        }

        _disposedValue = true;
    }

    ~OfficeFileDialogFilters()
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

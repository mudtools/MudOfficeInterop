//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imp;

/// <summary>
/// Office SelectedItems 集合对象的二次封装实现类
/// 实现 IOfficeSelectedItems 接口
/// </summary>
internal class OfficeSelectedItems : IOfficeSelectedItems
{
    private MsCore.FileDialogSelectedItems _selectedItems;
    private bool _disposedValue = false;

    internal OfficeSelectedItems(MsCore.FileDialogSelectedItems selectedItems)
    {
        _selectedItems = selectedItems ?? throw new ArgumentNullException(nameof(selectedItems));
    }

    #region 基础属性
    public int Count => _selectedItems.Count;

    public string this[int index] => _selectedItems.Item(index);

    public object Parent => _selectedItems.Parent;

    public object Application => _selectedItems.Application;
    #endregion    

    #region 高级功能
    public string[] GetAllItems()
    {
        var paths = new List<string>();
        for (int i = 1; i <= Count; i++)
        {
            paths.Add(this[i]);
        }
        return paths.ToArray();
    }
    #endregion

    #region IEnumerable<IOfficeSelectedItem> Support
    public IEnumerator<string> GetEnumerator()
    {
        for (int i = 1; i <= _selectedItems.Count; i++)
        {
            yield return _selectedItems.Item(i);
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
            try
            {
                // 释放底层COM对象
                if (_selectedItems != null)
                    Marshal.ReleaseComObject(_selectedItems);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _selectedItems = null;
        }

        _disposedValue = true;
    }

    ~OfficeSelectedItems()
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
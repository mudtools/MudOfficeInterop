
namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// IWordDropDown 接口的内部实现类，包装 Microsoft.Office.Interop.Word.DropDown COM 对象。
/// </summary>
internal class WordDropDown : IWordDropDown
{
    private MsWord.DropDown _dropDown;
    private bool _disposedValue;

    internal WordDropDown(MsWord.DropDown dropDown)
    {
        _dropDown = dropDown ?? throw new ArgumentNullException(nameof(dropDown));
        _disposedValue = false;
    }

    #region IWordDropDown 属性实现

    /// <summary>
    /// 获取或设置当前选中项的索引（从1开始）。
    /// </summary>
    public int Value
    {
        get => _dropDown?.Value ?? 0;
        set => _dropDown.Value = value;
    }

    /// <summary>
    /// 获取或设置默认选中项的索引（从1开始）。
    /// </summary>
    public int Default
    {
        get => _dropDown?.Default ?? 0;
        set => _dropDown.Default = value;
    }

    /// <summary>
    /// 获取下拉列表中的所有选项。
    /// </summary>
    public IReadOnlyList<string> ListEntries
    {
        get
        {
            if (_dropDown?.ListEntries == null) return new List<string>();

            var entries = new List<string>();
            var comEntries = _dropDown.ListEntries;
            try
            {
                for (int i = 1; i <= comEntries.Count; i++)
                {
                    entries.Add(comEntries[i].Name);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(comEntries);
            }
            return entries;
        }
    }

    /// <summary>
    /// 获取该下拉列表对象是否有效。
    /// </summary>
    public bool Valid => _dropDown?.Valid ?? false;

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _dropDown != null)
        {
            Marshal.ReleaseComObject(_dropDown);
            _dropDown = null;
        }
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
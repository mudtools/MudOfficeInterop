
namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// IWordCheckBox 接口的内部实现类，包装 Microsoft.Office.Interop.Word.CheckBox COM 对象。
/// </summary>
internal class WordCheckBox : IWordCheckBox
{
    private MsWord.CheckBox _checkBox;
    private bool _disposedValue;

    internal WordCheckBox(MsWord.CheckBox checkBox)
    {
        _checkBox = checkBox ?? throw new ArgumentNullException(nameof(checkBox));
        _disposedValue = false;
    }

    #region IWordCheckBox 属性实现

    /// <summary>
    /// 获取或设置复选框的当前值（选中状态）。
    /// </summary>
    public bool Value
    {
        get => _checkBox?.Value ?? false;
        set => _checkBox.Value = value;
    }

    /// <summary>
    /// 获取或设置复选框的默认值。
    /// </summary>
    public bool Default
    {
        get => _checkBox?.Default ?? false;
        set => _checkBox.Default = value;
    }

    /// <summary>
    /// 获取或设置是否自动调整大小。
    /// </summary>
    public bool AutoSize
    {
        get => _checkBox?.AutoSize ?? false;
        set => _checkBox.AutoSize = value;
    }

    /// <summary>
    /// 获取该复选框对象是否有效。
    /// </summary>
    public bool Valid => _checkBox?.Valid ?? false;

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _checkBox != null)
        {
            Marshal.ReleaseComObject(_checkBox);
            _checkBox = null;
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
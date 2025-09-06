//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// 对 Microsoft.Office.Core.PickerResult 的二次封装实现类。
/// 提供安全访问选取器结果属性的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficePickerResult : IOfficePickerResult
{
    private MsCore.PickerResult _pickerResult;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 PickerResult 对象。
    /// </summary>
    /// <param name="pickerResult">原始的 COM PickerResult 对象。</param>
    internal OfficePickerResult(MsCore.PickerResult pickerResult)
    {
        _pickerResult = pickerResult ?? throw new ArgumentNullException(nameof(pickerResult));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public string Id
    {
        get => _pickerResult?.Id ?? string.Empty;
    }

    /// <inheritdoc/>
    public string DisplayName
    {
        get => _pickerResult?.DisplayName ?? string.Empty;
        set
        {
            if (_pickerResult != null)
                _pickerResult.DisplayName = value;
        }
    }

    /// <inheritdoc/>
    public string Type
    {
        get => _pickerResult?.Type ?? string.Empty;
        set
        {
            if (_pickerResult != null)
                _pickerResult.Type = value;
        }
    }

    /// <inheritdoc/>
    public string SIPId
    {
        get => _pickerResult?.SIPId ?? string.Empty;
        set
        {
            if (_pickerResult != null)
                _pickerResult.SIPId = value;
        }
    }

    /// <inheritdoc/>
    public object? SubItems
    {
        get => _pickerResult?.SubItems ?? null;
        set
        {
            if (_pickerResult != null)
                _pickerResult.SubItems = value;
        }
    }

    /// <inheritdoc/>
    public object? DuplicateResults
    {
        get => _pickerResult?.DuplicateResults;
    }

    /// <inheritdoc/>
    public object? ItemData
    {
        get => _pickerResult?.ItemData ?? null;
        set
        {
            if (_pickerResult != null)
                _pickerResult.ItemData = value;
        }
    }

    /// <inheritdoc/>
    public IOfficePickerFields? Fields
    {
        get => _pickerResult?.Fields != null ? new OfficePickerFields(_pickerResult.Fields) : null;
        set
        {
            if (_pickerResult != null && value != null)
                _pickerResult.Fields = ((OfficePickerFields)value)._pickerFields;
        }
    }
    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放资源的核心方法。
    /// </summary>
    /// <param name="disposing">是否由 Dispose() 调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _pickerResult != null)
        {
            try
            {
                Marshal.ReleaseComObject(_pickerResult);
            }
            catch
            {
                // 忽略释放异常
            }
            _pickerResult = null;
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
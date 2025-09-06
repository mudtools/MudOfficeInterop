//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.PickerProperties 的二次封装实现类。
/// 提供安全访问选取器属性集合的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficePickerProperties : IOfficePickerProperties
{
    private MsCore.PickerProperties _pickerProperties;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 PickerProperties 对象。
    /// </summary>
    /// <param name="pickerProperties">原始的 COM PickerProperties 对象。</param>
    internal OfficePickerProperties(MsCore.PickerProperties pickerProperties)
    {
        _pickerProperties = pickerProperties ?? throw new ArgumentNullException(nameof(pickerProperties));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Count => _pickerProperties?.Count ?? 0;

    /// <inheritdoc/>
    public IOfficePickerProperty this[int index]
    {
        get
        {
            if (_pickerProperties == null || index < 1 || index > Count)
                return null;

            try
            {
                var property = _pickerProperties[index];
                return property != null ? new OfficePickerProperty(property) : null;
            }
            catch
            {
                return null;
            }
        }
    }
    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IOfficePickerProperty Add(string id, string value, MsoPickerField format = MsoPickerField.msoPickerFieldUnknown)
    {
        if (_pickerProperties == null || string.IsNullOrEmpty(id))
            return null;

        try
        {
            var property = _pickerProperties.Add(id, value, (MsCore.MsoPickerField)(int)format);
            return property != null ? new OfficePickerProperty(property) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public void Remove(string id)
    {
        if (_pickerProperties == null || string.IsNullOrEmpty(id))
            return;

        _pickerProperties.Remove(id);
    }


    /// <inheritdoc/>
    public bool Contains(string id)
    {
        if (_pickerProperties == null || string.IsNullOrEmpty(id))
            return false;

        // 遍历查找匹配键的属性
        for (int i = 1; i <= Count; i++)
        {
            var property = _pickerProperties[i];
            if (property != null && string.Equals(property.Id, id, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }
        return false;
    }
    #endregion

    #region IEnumerable<IOfficePickerProperty> 实现

    /// <inheritdoc/>
    public IEnumerator<IOfficePickerProperty> GetEnumerator()
    {
        if (_pickerProperties == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var property = _pickerProperties[i];
            if (property != null)
                yield return new OfficePickerProperty(property);
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
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

        if (disposing && _pickerProperties != null)
        {
            try
            {
                Marshal.ReleaseComObject(_pickerProperties);
            }
            catch
            {
                // 忽略释放异常
            }
            _pickerProperties = null;
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

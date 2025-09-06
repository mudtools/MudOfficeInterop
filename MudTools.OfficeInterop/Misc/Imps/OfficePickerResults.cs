//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.PickerResults 的二次封装实现类。
/// 提供安全访问选取器结果集合的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficePickerResults : IOfficePickerResults
{
    private static readonly ILog log = LogManager.GetLogger(typeof(OfficePropertyTests));
    internal MsCore.PickerResults _pickerResults;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 PickerResults 对象。
    /// </summary>
    /// <param name="pickerResults">原始的 COM PickerResults 对象。</param>
    internal OfficePickerResults(MsCore.PickerResults pickerResults)
    {
        _pickerResults = pickerResults ?? throw new ArgumentNullException(nameof(pickerResults));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Count => _pickerResults?.Count ?? 0;

    /// <inheritdoc/>
    public IOfficePickerResult this[int index]
    {
        get
        {
            if (_pickerResults == null || index < 1 || index > Count)
                return null;

            try
            {
                var result = _pickerResults[index];
                return result != null ? new OfficePickerResult(result) : null;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }
    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IOfficePickerResult Add(string id, string name, string type)
    {
        if (_pickerResults == null)
            return null;

        try
        {
            var result = _pickerResults.Add(id, name, type);
            return result != null ? new OfficePickerResult(result) : null;
        }
        catch (COMException ce)
        {
            log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
            return null;
        }
    }
    #endregion

    #region IEnumerable<IOfficePickerResult> 实现

    /// <inheritdoc/>
    public IEnumerator<IOfficePickerResult> GetEnumerator()
    {
        if (_pickerResults == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var result = _pickerResults[i];
            if (result != null)
                yield return new OfficePickerResult(result);
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

        if (disposing && _pickerResults != null)
        {
            try
            {
                Marshal.ReleaseComObject(_pickerResults);
            }
            catch
            {
                // 忽略释放异常
            }
            _pickerResults = null;
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
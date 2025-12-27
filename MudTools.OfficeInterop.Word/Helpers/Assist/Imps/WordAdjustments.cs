//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 形状调整点对象的封装实现类。
/// </summary>
internal class WordAdjustments : IWordAdjustments
{
    private MsWord.Adjustments _adjustments;
    private bool _disposedValue;

    internal WordAdjustments(MsWord.Adjustments adjustments)
    {
        _adjustments = adjustments ?? throw new ArgumentNullException(nameof(adjustments));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _adjustments != null ? new WordApplication(_adjustments.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _adjustments?.Parent;

    /// <inheritdoc/>
    public int Count => _adjustments?.Count ?? 0;

    /// <inheritdoc/>
    public float this[int index]
    {
        get
        {
            if (_adjustments == null || index < 1 || index > Count)
                return 0f;

            try
            {
                return _adjustments[index];
            }
            catch (COMException)
            {
                return 0f;
            }
        }
        set
        {
            if (_adjustments == null || index < 1 || index > Count)
                return;

            try
            {
                _adjustments[index] = value;
            }
            catch (COMException)
            {
                // 静默忽略调整失败
            }
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _adjustments != null)
        {
            Marshal.ReleaseComObject(_adjustments);
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
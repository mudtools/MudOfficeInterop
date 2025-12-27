//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 校对错误集合的封装实现类。
/// </summary>
internal class WordProofreadingErrors : IWordProofreadingErrors
{
    private MsWord.ProofreadingErrors _proofreadingErrors;
    private bool _disposedValue;

    internal WordProofreadingErrors(MsWord.ProofreadingErrors proofreadingErrors)
    {
        _proofreadingErrors = proofreadingErrors ?? throw new ArgumentNullException(nameof(proofreadingErrors));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _proofreadingErrors != null ? new WordApplication(_proofreadingErrors.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _proofreadingErrors?.Parent;

    /// <inheritdoc/>
    public int Count => _proofreadingErrors?.Count ?? 0;

    /// <inheritdoc/>
    public IWordRange this[int index]
    {
        get
        {
            if (index < 1 || index > Count || _proofreadingErrors == null) return null;
            try
            {
                var comRange = _proofreadingErrors[index];
                return comRange != null ? new WordRange(comRange) : null;
            }
            catch (COMException)
            {
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public WdProofreadingErrorType Type => _proofreadingErrors?.Type != null ? (WdProofreadingErrorType)(int)_proofreadingErrors?.Type : WdProofreadingErrorType.wdSpellingError;

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _proofreadingErrors != null)
        {
            Marshal.ReleaseComObject(_proofreadingErrors);
            _proofreadingErrors = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable 实现

    public IEnumerator<IWordRange> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
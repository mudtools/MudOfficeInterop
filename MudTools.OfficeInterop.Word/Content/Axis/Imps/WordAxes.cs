//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Axes 的封装实现类。
/// </summary>
internal class WordAxes : IWordAxes
{
    private MsWord.Axes _axes;
    private bool _disposedValue;

    internal WordAxes(MsWord.Axes axes)
    {
        _axes = axes ?? throw new ArgumentNullException(nameof(axes));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _axes != null ? new WordApplication(_axes.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _axes?.Parent;

    /// <inheritdoc/>
    public int Count => _axes?.Count ?? 0;

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordAxis this[XlAxisType type, XlAxisGroup axisGroup]
    {
        get
        {
            var comAxis = _axes[(MsWord.XlAxisType)(int)type, (MsWord.XlAxisGroup)(int)axisGroup];
            return new WordAxis(comAxis);
        }
    }

    #endregion

    #region 方法实现
    /// <inheritdoc/>
    public IWordAxis CategoryAxis => this[XlAxisType.xlCategory, XlAxisGroup.xlPrimary];

    /// <inheritdoc/>
    public IWordAxis ValueAxis => this[XlAxisType.xlValue, XlAxisGroup.xlPrimary];

    /// <inheritdoc/>
    public IWordAxis SecondaryCategoryAxis => this[XlAxisType.xlCategory, XlAxisGroup.xlSecondary];

    /// <inheritdoc/>
    public IWordAxis SecondaryValueAxis => this[XlAxisType.xlValue, XlAxisGroup.xlSecondary];

    /// <inheritdoc/>
    public List<XlAxisType> GetAxisTypes()
    {
        var types = new List<XlAxisType>();
        foreach (var axis in this)
        {
            if (axis?.Type != null && !types.Contains(axis.Type))
                types.Add(axis.Type);
        }
        return types;
    }

    /// <inheritdoc/>
    public int CountByType(XlAxisType type)
    {
        int count = 0;
        foreach (var axis in this)
        {
            if (axis?.Type == type)
                count++;
        }
        return count;
    }

    #endregion

    #region 枚举支持

    /// <inheritdoc/>
    public IEnumerator<IWordAxis> GetEnumerator()
    {
        foreach (var axis in _axes)
        {
            yield return new WordAxis(axis as MsWord.Axis);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _axes != null)
        {
            Marshal.ReleaseComObject(_axes);
            _axes = null;
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
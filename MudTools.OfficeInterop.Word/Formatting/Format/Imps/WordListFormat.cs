//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.ListFormat 的封装实现类。
/// </summary>
internal class WordListFormat : IWordListFormat
{
    private MsWord.ListFormat _listFormat;
    private bool _disposedValue;

    internal WordListFormat(MsWord.ListFormat listFormat)
    {
        _listFormat = listFormat ?? throw new ArgumentNullException(nameof(listFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _listFormat != null ? new WordApplication(_listFormat.Application) : null;

    /// <inheritdoc/>
    public object Parent => _listFormat?.Parent;

    /// <inheritdoc/>
    public int ListLevelNumber
    {
        get => _listFormat?.ListLevelNumber ?? 0;
        set
        {
            if (_listFormat != null)
                _listFormat.ListLevelNumber = value;
        }
    }

    /// <inheritdoc/>
    public IWordListTemplate? ListTemplate
    {
        get => _listFormat?.ListTemplate != null ? new WordListTemplate(_listFormat.ListTemplate) : null;
    }


    /// <inheritdoc/>
    public WdListType ListType => _listFormat?.ListType != null ? (WdListType)(int)_listFormat?.ListType : WdListType.wdListMixedNumbering;


    /// <inheritdoc/>
    public string ListString => _listFormat?.ListString ?? string.Empty;


    /// <inheritdoc/>
    public bool SingleList
    {
        get => _listFormat?.SingleList ?? false;
    }

    /// <inheritdoc/>
    public bool SingleListTemplate
    {
        get => _listFormat?.SingleListTemplate ?? false;
    }
    #endregion

    #region 方法实现

    public void ApplyBulletDefault(WdDefaultListBehavior? DefaultListBehavior = null)
    {
        _listFormat?.ApplyBulletDefault(
            DefaultListBehavior.ComArgsConvert(d => d.EnumConvert(MsWord.WdDefaultListBehavior.wdWord10ListBehavior)));
    }

    /// <inheritdoc/>
    public void ApplyNumberDefault(WdDefaultListBehavior? DefaultListBehavior = null)
    {
        _listFormat?.ApplyNumberDefault(
            DefaultListBehavior.ComArgsConvert(d => d.EnumConvert(MsWord.WdDefaultListBehavior.wdWord10ListBehavior)));
    }

    /// <inheritdoc/>
    public void ApplyOutlineNumberDefault(WdDefaultListBehavior? DefaultListBehavior = null)
    {
        _listFormat?.ApplyOutlineNumberDefault(
                    DefaultListBehavior.ComArgsConvert(d => d.EnumConvert(MsWord.WdDefaultListBehavior.wdWord10ListBehavior)));
    }

    /// <inheritdoc/>
    public void ApplyListTemplateWithLevel(
        IWordListTemplate listTemplate,
        bool continuePreviousList,
        WdListApplyTo applyTo,
        WdDefaultListBehavior defaultListBehavior)
    {
        _listFormat?.ApplyListTemplateWithLevel(
            listTemplate is WordListTemplate template ? template._listTemplate : null,
            continuePreviousList,
            applyTo.EnumConvert(MsWord.WdListApplyTo.wdListApplyToSelection),
            defaultListBehavior.EnumConvert(MsWord.WdDefaultListBehavior.wdWord10ListBehavior));
    }

    public void ApplyListTemplate(
        IWordListTemplate listTemplate,
        bool continuePreviousList,
        WdListApplyTo applyTo,
        WdDefaultListBehavior defaultListBehavior)
    {
        _listFormat?.ApplyListTemplate(
             listTemplate is WordListTemplate template ? template._listTemplate : null,
             continuePreviousList,
             applyTo.EnumConvert(MsWord.WdListApplyTo.wdListApplyToSelection),
             defaultListBehavior.EnumConvert(MsWord.WdDefaultListBehavior.wdWord10ListBehavior));
    }

    /// <inheritdoc/>
    public void RemoveNumbers()
    {
        _listFormat?.RemoveNumbers();
    }

    /// <inheritdoc/>
    public WdContinue CanContinuePreviousList(IWordListTemplate listTemplate)
    {
        if (_listFormat == null)
            return WdContinue.wdContinueDisabled;
        var r = _listFormat?.CanContinuePreviousList(listTemplate is WordListTemplate template ? template._listTemplate : null);
        return r != null ? r.EnumConvert(WdContinue.wdContinueDisabled) : WdContinue.wdContinueDisabled;
    }

    /// <inheritdoc/>
    public void ListIndent()
    {
        _listFormat?.ListIndent();
    }

    /// <inheritdoc/>
    public void ListOutdent()
    {
        _listFormat?.ListOutdent();
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _listFormat != null)
        {
            Marshal.ReleaseComObject(_listFormat);
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
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Interop.Word;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Field 的实现类。
/// </summary>
internal class WordField : IWordField
{
    private MsWord.Field _field;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="field">原始 COM Field 对象。</param>
    internal WordField(MsWord.Field field)
    {
        _field = field ?? throw new ArgumentNullException(nameof(field));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _field != null ? new WordApplication(_field.Application) : null;

    /// <inheritdoc/>
    public object Parent => _field?.Parent;

    /// <inheritdoc/>
    public WdFieldType Type => _field?.Type != null ? (WdFieldType)(int)_field?.Type : WdFieldType.wdFieldNoteRef;


    /// <inheritdoc/>
    public WdFieldKind Kind => _field?.Kind != null ? (WdFieldKind)(int)_field?.Kind : WdFieldKind.wdFieldKindNone;


    /// <inheritdoc/>
    public IWordRange? ResultRange =>
        _field?.Result != null ? new WordRange(_field.Result) : null;

    /// <inheritdoc/>
    public IWordRange? CodeRange =>
        _field?.Code != null ? new WordRange(_field.Code) : null;

    /// <inheritdoc/>
    public object Parent => _field?.Parent;

    /// <inheritdoc/>
    public bool Locked
    {
        get => _field?.Locked != null && _field.Locked;
        set
        {
            if (_field != null)
                _field.Locked = value;
        }
    }

    /// <inheritdoc/>
    public int Index => _field?.Index ?? 0;

    /// <inheritdoc/>
    public string Data
    {
        get => _field?.Data ?? string.Empty;
        set
        {
            if (_field != null)
                _field.Data = value;
        }
    }

    /// <inheritdoc/>
    public string Result
    {
        get => _field?.Result?.Text ?? string.Empty;
        set
        {
            if (_field?.Result != null)
                _field.Result.Text = value;
        }
    }

    /// <inheritdoc/>
    public string Code
    {
        get => _field?.Code?.Text ?? string.Empty;
        set
        {
            if (_field?.Code != null)
                _field.Code.Text = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowCodes
    {
        get => _field?.ShowCodes != null && _field.ShowCodes;
        set
        {
            if (_field != null)
                _field.ShowCodes = value;
        }
    }

    /// <inheritdoc/>
    public IWordField? NextField =>
        _field?.Next != null ? new WordField(_field.Next) : null;

    /// <inheritdoc/>
    public IWordField? PreviousField =>
        _field?.Previous != null ? new WordField(_field.Previous) : null;

    /// <inheritdoc/>
    public bool IsLinked => LinkFormat != null;

    /// <inheritdoc/>
    public IWordLinkFormat? LinkFormat =>
         _field?.LinkFormat != null ? new WordLinkFormat(_field.LinkFormat) : null;

    /// <inheritdoc/>
    public IWordOLEFormat? OLEFormat =>
         _field?.OLEFormat != null ? new WordOLEFormat(_field.OLEFormat) : null;

    /// <inheritdoc/>
    public bool IsDateField => Type == WdFieldType.wdFieldDate ||
                              Type == WdFieldType.wdFieldCreateDate ||
                              Type == WdFieldType.wdFieldPrintDate;

    /// <inheritdoc/>
    public bool IsPageField => Type == WdFieldType.wdFieldPage ||
                              Type == WdFieldType.wdFieldNumPages;

    /// <inheritdoc/>
    public bool IsTOCField => Type == WdFieldType.wdFieldTOC ||
                             Type == WdFieldType.wdFieldTOA ||
                             Type == WdFieldType.wdFieldIndex;
    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public bool Update()
    {
        if (_field == null)
            return false;

        return _field.Update();
    }

    /// <inheritdoc/>
    public void Unlink()
    {
        _field?.Unlink();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _field?.Delete();
    }

    /// <inheritdoc/>
    public void Select()
    {
        _field?.Select();
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _field?.Copy();
    }

    /// <inheritdoc/>
    public void Cut()
    {
        _field?.Cut();
    }


    /// <inheritdoc/>
    public void DoClick()
    {
        _field?.DoClick();
    }

    /// <inheritdoc/>
    public bool ValidateCode()
    {
        if (_field == null)
            return false;

        try
        {
            return Update();
        }
        catch
        {
            return false;
        }
    }


    /// <inheritdoc/>
    public string GetSourcePath()
    {
        if (_field == null)
            return string.Empty;

        try
        {
            return LinkFormat?.SourceFullName ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    /// <inheritdoc/>
    public void SetCode(string code)
    {
        if (_field?.Code != null && !string.IsNullOrEmpty(code))
        {
            _field.Code.Text = code;
        }
    }

    /// <inheritdoc/>
    public void SetResult(string result)
    {
        if (_field?.Result != null && !string.IsNullOrEmpty(result))
        {
            _field.Result.Text = result;
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_field?.Next != null)
            {
                Marshal.ReleaseComObject(_field.Next);
            }
            if (_field?.Previous != null)
            {
                Marshal.ReleaseComObject(_field.Previous);
            }
            if (_field?.Result != null)
            {
                Marshal.ReleaseComObject(_field.Result);
            }
            if (_field?.Code != null)
            {
                Marshal.ReleaseComObject(_field.Code);
            }
            // 释放域对象本身
            if (_field != null)
            {
                Marshal.ReleaseComObject(_field);
                _field = null;
            }
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
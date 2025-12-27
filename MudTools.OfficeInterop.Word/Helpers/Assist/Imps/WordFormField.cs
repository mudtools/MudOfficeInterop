//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表单域对象的封装实现类。
/// </summary>
internal class WordFormField : IWordFormField
{
    private MsWord.FormField _formField;
    private bool _disposedValue;

    internal WordFormField(MsWord.FormField formField)
    {
        _formField = formField ?? throw new ArgumentNullException(nameof(formField));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _formField != null ? new WordApplication(_formField.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _formField?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _formField?.Name ?? string.Empty;
        set
        {
            if (_formField != null)
                _formField.Name = value;
        }
    }

    /// <inheritdoc/>
    public string Result
    {
        get => _formField?.Result ?? string.Empty;
        set
        {
            if (_formField != null)
                _formField.Result = value;
        }
    }

    /// <inheritdoc/>
    public WdFieldType Type => _formField?.Type != null ? (WdFieldType)(int)_formField?.Type : WdFieldType.wdFieldAutoText;


    public IWordTextInput? TextInput
    {
        get
        {
            if (_formField?.TextInput != null)
                return new WordTextInput(_formField.TextInput);
            return null;
        }
    }

    public IWordCheckBox? CheckBox
    {
        get
        {
            if (_formField?.CheckBox != null)
                return new WordCheckBox(_formField.CheckBox);
            return null;
        }
    }

    public IWordDropDown? DropDown
    {
        get
        {
            if (_formField?.DropDown != null)
                return new WordDropDown(_formField.DropDown);
            return null;
        }
    }

    public IWordFormField? Next
    {
        get
        {
            if (_formField?.Next != null)
                return new WordFormField(_formField.Next);
            return null;
        }
    }

    public IWordFormField? Previous
    {
        get
        {
            if (_formField?.Previous != null)
                return new WordFormField(_formField.Previous);
            return null;
        }
    }

    public IWordRange? Range
    {
        get
        {
            if (_formField?.Range != null)
                return new WordRange(_formField.Range);
            return null;
        }
    }

    /// <inheritdoc/>
    public bool CheckBox_Checked
    {
        get => _formField?.CheckBox?.Valid ?? false ? _formField.CheckBox.Value : false;
        set
        {
            if (_formField?.CheckBox?.Valid == true)
                _formField.CheckBox.Value = value;
        }
    }

    /// <inheritdoc/>
    public string TextInput_Default
    {
        get => _formField?.TextInput?.Default ?? string.Empty;
        set
        {
            if (_formField?.TextInput != null)
                _formField.TextInput.Default = value;
        }
    }

    /// <inheritdoc/>
    public int DropDown_Default
    {
        get => _formField?.DropDown?.Default ?? 0;
        set
        {
            if (_formField?.DropDown != null)
                _formField.DropDown.Default = value;
        }
    }

    /// <inheritdoc/>
    public List<string> DropDown_ListEntries
    {
        get
        {
            var list = new List<string>();
            if (_formField?.DropDown?.ListEntries != null)
            {
                foreach (MsWord.ListEntry entry in _formField.DropDown.ListEntries)
                {
                    list.Add(entry.Name);
                }
            }
            return list;
        }
    }

    #endregion

    #region 方法实现

    public void Copy()
    {
        _formField?.Copy();
    }

    public void Cut()
    {
        _formField?.Cut();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _formField?.Delete();
    }

    /// <inheritdoc/>
    public void Select()
    {
        _formField?.Select();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _formField != null)
        {
            Marshal.ReleaseComObject(_formField);
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
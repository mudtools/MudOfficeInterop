//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// IWordTextInput 接口的内部实现类，包装 Microsoft.Office.Interop.Word.TextInput COM 对象。
/// </summary>
internal class WordTextInput : IWordTextInput
{
    private MsWord.TextInput _textInput;
    private bool _disposedValue;

    internal WordTextInput(MsWord.TextInput textInput)
    {
        _textInput = textInput ?? throw new ArgumentNullException(nameof(textInput));
        _disposedValue = false;
    }

    #region IWordTextInput 属性实现

    /// <summary>
    /// 获取或设置默认文本。
    /// </summary>
    public string Default
    {
        get => _textInput?.Default;
        set => _textInput.Default = value;
    }

    /// <summary>
    /// 获取文本格式。
    /// </summary>
    public string Format => _textInput?.Format;

    /// <summary>
    /// 获取文本域类型。
    /// </summary>
    public WdTextFormFieldType Type => _textInput?.Type.EnumConvert(WdTextFormFieldType.wdRegularText) ?? WdTextFormFieldType.wdRegularText;

    #endregion

    public void Clear()
    {
        _textInput?.Clear();
    }

    public void EditType(WdTextFormFieldType type, string? @default = null, string? format = null, bool? enabled = null)
    {
        var comType = type.EnumConvert(MsWord.WdTextFormFieldType.wdRegularText);
        _textInput?.EditType(comType, @default.ComArgsVal(), format.ComArgsVal(), enabled.ComArgsVal());
    }

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _textInput != null)
        {
            Marshal.ReleaseComObject(_textInput);
            _textInput = null;
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
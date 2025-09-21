//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;


/// <summary>
/// TextEffectFormat COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class ExcelTextEffectFormat : IExcelTextEffectFormat
{
    /// <summary>
    /// 内部持有的原始 COM 对象。
    /// </summary>
    internal MsExcel.TextEffectFormat _textEffectFormat;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="textEffectFormat">原始的 TextEffectFormat COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 textEffectFormat 为 null 时抛出。</exception>
    internal ExcelTextEffectFormat(MsExcel.TextEffectFormat textEffectFormat)
    {
        _textEffectFormat = textEffectFormat ?? throw new ArgumentNullException(nameof(textEffectFormat));
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的受保护虚方法，支持派生类重写。
    /// </summary>
    /// <param name="disposing">是否由用户代码显式调用释放。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放托管资源：释放 COM 对象
            if (_textEffectFormat != null)
            {
                Marshal.ReleaseComObject(_textEffectFormat);
                _textEffectFormat = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 公开的 Dispose 方法，用于显式释放资源。
    /// 调用后对象不应再被使用。
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取此对象的父对象（通常是 Shape）。
    /// </summary>
    public object Parent => _textEffectFormat?.Parent;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication Application =>
        _textEffectFormat?.Application != null
            ? new ExcelApplication(_textEffectFormat.Application as MsExcel.Application)
            : null;

    /// <summary>
    /// 获取或设置艺术字的对齐方式（左对齐、居中、右对齐等）。
    /// 默认值：msoTextEffectAlignmentMixed
    /// </summary>
    public MsoTextEffectAlignment Alignment
    {
        get => _textEffectFormat != null
            ? _textEffectFormat.Alignment.EnumConvert(MsoTextEffectAlignment.msoTextEffectAlignmentMixed)
            : MsoTextEffectAlignment.msoTextEffectAlignmentMixed;

        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.Alignment = value.EnumConvert(MsCore.MsoTextEffectAlignment.msoTextEffectAlignmentMixed);
        }
    }

    public bool FontBold
    {
        get => _textEffectFormat != null && _textEffectFormat.FontBold.ConvertToBool();
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.FontBold = value.ConvertTriState();
        }
    }

    public bool FontItalic
    {
        get => _textEffectFormat != null && _textEffectFormat.FontItalic.ConvertToBool();
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.FontItalic = value.ConvertTriState();
        }
    }

    public bool NormalizedHeight
    {
        get => _textEffectFormat != null && _textEffectFormat.NormalizedHeight.ConvertToBool();
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.NormalizedHeight = value.ConvertTriState();
        }
    }

    public string FontName
    {
        get => _textEffectFormat != null ? _textEffectFormat.FontName : string.Empty;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.FontName = value;
        }
    }


    /// <summary>
    /// 获取或设置艺术字的字体大小缩放比例（相对于原始设计尺寸）。
    /// 1.0 = 100%，2.0 = 200%。
    /// </summary>
    public float FontSize
    {
        get => _textEffectFormat != null ? _textEffectFormat.FontSize : 0f;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.FontSize = value;
        }
    }

    /// <summary>
    /// 获取或设置艺术字是否启用字偶距调整（Kerning），以优化字符间距。
    /// </summary>
    public bool KernedPairs
    {
        get => _textEffectFormat != null && _textEffectFormat.KernedPairs.ConvertToBool();
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.KernedPairs = value.ConvertTriState();
        }
    }

    public bool RotatedChars
    {
        get => _textEffectFormat != null && _textEffectFormat.RotatedChars.ConvertToBool();
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.RotatedChars = value.ConvertTriState();
        }
    }

    public float Tracking
    {
        get => _textEffectFormat != null ? _textEffectFormat.Tracking : 0f;
        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.Tracking = value;
        }
    }


    /// <summary>
    /// 获取或设置艺术字的预设文本路径样式（如拱形、波浪形等）。
    /// 默认值：msoTextEffectShapeMixed
    /// 注意：实际 COM 属性名为 .PresetTextEffectShape
    /// </summary>
    public MsoPresetTextEffectShape PresetShape
    {
        get => _textEffectFormat != null
            ? _textEffectFormat.PresetShape.EnumConvert(MsoPresetTextEffectShape.msoTextEffectShapeMixed)
            : MsoPresetTextEffectShape.msoTextEffectShapeMixed;

        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.PresetShape = value.EnumConvert(MsCore.MsoPresetTextEffectShape.msoTextEffectShapeMixed);
        }
    }

    public MsoPresetTextEffect PresetTextEffect
    {
        get => _textEffectFormat != null
            ? _textEffectFormat.PresetTextEffect.EnumConvert(MsoPresetTextEffect.msoTextEffectMixed)
            : MsoPresetTextEffect.msoTextEffectMixed;

        set
        {
            if (_textEffectFormat != null)
                _textEffectFormat.PresetTextEffect = value.EnumConvert(MsCore.MsoPresetTextEffect.msoTextEffectMixed);
        }
    }

    /// <summary>
    /// 获取或设置艺术字文本内容。
    /// </summary>
    public string Text
    {
        get => _textEffectFormat?.Text ?? string.Empty;
        set
        {
            if (_textEffectFormat != null && value != null)
                _textEffectFormat.Text = value;
        }
    }

    /// <summary>
    /// 将艺术字文本方向切换为垂直（Toggle）。
    /// 调用一次垂直，再调用一次恢复水平。
    /// </summary>
    public void ToggleVerticalText()
    {
        _textEffectFormat?.ToggleVerticalText();
    }


}
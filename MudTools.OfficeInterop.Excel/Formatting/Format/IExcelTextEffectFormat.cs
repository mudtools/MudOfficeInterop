//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// 表示 Excel 艺术字（WordArt）文本效果格式的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.TextEffectFormat
/// 用于控制艺术字的文字方向、对齐、缩放、路径样式等。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelTextEffectFormat : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Shape）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置艺术字的对齐方式（左对齐、居中、右对齐等）。
    /// 使用 <see cref="MsoTextEffectAlignment"/> 枚举。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoTextEffectAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置艺术字的预设文本效果样式。
    /// 该属性值对应于艺术字库对话框中从左到右、从上到下排列的格式。
    /// 设置此属性会自动设置指定形状的许多其他格式属性。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetTextEffect PresetTextEffect { get; set; }

    /// <summary>
    /// 获取或设置艺术字的字体名称（如 Arial、Times New Roman 等）。
    /// </summary>
    string FontName { get; set; }

    /// <summary>
    /// 获取或设置艺术字是否使用粗体样式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool FontBold { get; set; }

    /// <summary>
    /// 获取或设置艺术字是否使用斜体样式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool FontItalic { get; set; }

    /// <summary>
    /// 获取或设置是否将所有字符调整为相同高度（规范化高度）。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool NormalizedHeight { get; set; }

    /// <summary>
    /// 获取或设置字符是否沿路径旋转（适用于弯曲或倾斜的艺术字效果）。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool RotatedChars { get; set; }

    /// <summary>
    /// 获取或设置字符间距（跟踪），控制字符之间的间距比例。
    /// 通常 0.0 表示无间距，正值增加间距，负值减少间距。
    /// </summary>
    float Tracking { get; set; }

    /// <summary>
    /// 获取或设置艺术字的字体大小缩放比例（相对于原始设计尺寸）。
    /// 1.0 = 100%，2.0 = 200%。
    /// </summary>
    float FontSize { get; set; }

    /// <summary>
    /// 获取或设置艺术字是否保持文本比例（锁定宽高比）。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool KernedPairs { get; set; }

    /// <summary>
    /// 获取或设置艺术字的预设文本路径样式（如拱形、波浪形等）。
    /// 使用 <see cref="MsoPresetTextEffectShape"/> 枚举。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetTextEffectShape PresetShape { get; set; }


    /// <summary>
    /// 获取或设置艺术字文本内容。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 将艺术字文本方向切换为垂直（Toggle）。
    /// 调用一次垂直，再调用一次恢复水平。
    /// </summary>
    void ToggleVerticalText();
}
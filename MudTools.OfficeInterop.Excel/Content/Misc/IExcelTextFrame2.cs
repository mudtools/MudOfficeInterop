//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示 Excel 中的文本框架对象，提供对文本框属性和行为的访问接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelTextFrame2 : IDisposable
{
    /// <summary>
    /// 获取对象的父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取对象所在的 Application 对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置文本框底部边距
    /// </summary>
    float MarginBottom { get; set; }

    /// <summary>
    /// 获取或设置文本框左侧边距
    /// </summary>
    float MarginLeft { get; set; }

    /// <summary>
    /// 获取或设置文本框右侧边距
    /// </summary>
    float MarginRight { get; set; }

    /// <summary>
    /// 获取或设置文本框顶部边距
    /// </summary>
    float MarginTop { get; set; }

    /// <summary>
    /// 获取或设置文本的方向
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoTextOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置文本水平锚点位置
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoHorizontalAnchor HorizontalAnchor { get; set; }

    /// <summary>
    /// 获取或设置文本垂直锚点位置
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoVerticalAnchor VerticalAnchor { get; set; }

    /// <summary>
    /// 获取或设置路径格式
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPathFormat PathFormat { get; set; }

    /// <summary>
    /// 获取或设置扭曲格式
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoWarpFormat WarpFormat { get; set; }

    /// <summary>
    /// 获取或设置文字艺术效果格式
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetTextEffect WordArtformat { get; set; }

    /// <summary>
    /// 获取或设置是否自动换行
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool WordWrap { get; set; }

    /// <summary>
    /// 获取或设置文本是否不随对象旋转
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool NoTextRotation { get; set; }

    /// <summary>
    /// 获取是否有文本内容
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasText { get; }

    /// <summary>
    /// 获取或设置自动调整大小模式
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoAutoSize AutoSize { get; set; }

    /// <summary>
    /// 获取三维格式设置对象
    /// </summary>
    IExcelThreeDFormat ThreeD { get; }

    /// <summary>
    /// 获取文本范围对象
    /// </summary>
    IOfficeTextRange2 TextRange { get; }

    /// <summary>
    /// 删除文本框中的所有文本
    /// </summary>
    void DeleteText();
}
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示Office段落格式的接口，提供对段落各种格式属性的访问和设置功能
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeParagraphFormat2 : IDisposable
{
    /// <summary>
    /// 获取段落的父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取段落的制表符停止点集合
    /// </summary>
    IOfficeTabStops2 TabStops { get; }

    /// <summary>
    /// 获取段落的项目符号格式设置
    /// </summary>
    IOfficeBulletFormat2 Bullet { get; }

    /// <summary>
    /// 获取或设置段落的对齐方式
    /// </summary>
    MsoParagraphAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置段落的基线对齐方式
    /// </summary>
    MsoBaselineAlignment BaselineAlignment { get; set; }

    /// <summary>
    /// 获取或设置远东语言的换行级别
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool FarEastLineBreakLevel { get; set; }

    /// <summary>
    /// 获取或设置首行缩进值
    /// </summary>
    float FirstLineIndent { get; set; }

    /// <summary>
    /// 获取或设置是否启用悬挂标点
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HangingPunctuation { get; set; }

    /// <summary>
    /// 获取或设置缩进级别
    /// </summary>
    int IndentLevel { get; set; }

    /// <summary>
    /// 获取或设置左缩进值
    /// </summary>
    float LeftIndent { get; set; }

    /// <summary>
    /// 获取或设置右缩进值
    /// </summary>
    float RightIndent { get; set; }

    /// <summary>
    /// 获取或设置段落后的间距
    /// </summary>
    float SpaceAfter { get; set; }

    /// <summary>
    /// 获取或设置段落前的间距
    /// </summary>
    float SpaceBefore { get; set; }

    /// <summary>
    /// 获取或设置段落内的行距
    /// </summary>
    float SpaceWithin { get; set; }

    /// <summary>
    /// 获取或设置文本方向
    /// </summary>
    MsoTextDirection TextDirection { get; set; }

    /// <summary>
    /// 获取或设置段落后间距的行距规则
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool LineRuleAfter { get; set; }

    /// <summary>
    /// 获取或设置段落前间距的行距规则
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool LineRuleBefore { get; set; }

    /// <summary>
    /// 获取或设置段落内行距的行距规则
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool LineRuleWithin { get; set; }

    /// <summary>
    /// 获取或设置是否启用自动换行
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool WordWrap { get; set; }
}
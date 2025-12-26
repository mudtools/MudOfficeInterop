//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中数学对象的边框框，用于控制数学公式周围的边框和删除线显示
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathBorderBox : IOfficeObject<IWordOMathBorderBox>, IDisposable
{
    /// <summary>
    /// 获取与此数学对象关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此数学对象的父级对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置边框框中包含的数学对象元素
    /// </summary>
    IWordOMath? E { get; }

    /// <summary>
    /// 获取或设置是否隐藏顶部边框
    /// </summary>
    bool HideTop { get; set; }

    /// <summary>
    /// 获取或设置是否隐藏底部边框
    /// </summary>
    bool HideBot { get; set; }

    /// <summary>
    /// 获取或设置是否隐藏左侧边框
    /// </summary>
    bool HideLeft { get; set; }

    /// <summary>
    /// 获取或设置是否隐藏右侧边框
    /// </summary>
    bool HideRight { get; set; }

    /// <summary>
    /// 获取或设置是否显示水平删除线
    /// </summary>
    bool StrikeH { get; set; }

    /// <summary>
    /// 获取或设置是否显示垂直删除线
    /// </summary>
    bool StrikeV { get; set; }

    /// <summary>
    /// 获取或设置是否显示从右下到左上(BLTR)的对角删除线
    /// </summary>
    bool StrikeBLTR { get; set; }

    /// <summary>
    /// 获取或设置是否显示从左上到右下(TLBR)的对角删除线
    /// </summary>
    bool StrikeTLBR { get; set; }
}
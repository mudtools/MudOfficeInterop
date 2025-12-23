//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的数学公式分隔符对象的接口
/// 提供对数学公式分隔符属性的访问和修改功能
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathDelim : IDisposable
{
    /// <summary>
    /// 获取与当前对象关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取当前对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置分隔符内的数学表达式参数
    /// </summary>
    IWordOMathArgs? E { get; }

    /// <summary>
    /// 获取或设置分隔符开始字符的Unicode值
    /// </summary>
    short BegChar { get; set; }

    /// <summary>
    /// 获取或设置分隔符分隔字符的Unicode值
    /// </summary>
    short SepChar { get; set; }

    /// <summary>
    /// 获取或设置分隔符结束字符的Unicode值
    /// </summary>
    short EndChar { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示分隔符是否随内容大小自动调整
    /// </summary>
    bool Grow { get; set; }

    /// <summary>
    /// 获取或设置分隔符的形状类型
    /// </summary>
    WdOMathShapeType Shape { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否不显示左侧字符
    /// </summary>
    bool NoLeftChar { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否不显示右侧字符
    /// </summary>
    bool NoRightChar { get; set; }
}
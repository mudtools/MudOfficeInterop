//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中某个词的同义词信息接口
/// 提供对Word应用程序中同义词、反义词、相关表达等相关信息的访问
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordSynonymInfo : IDisposable
{
    /// <summary>
    /// 获取与此对象关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此对象的父级对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个值，指示是否找到了当前词的同义词信息
    /// </summary>
    bool Found { get; }

    /// <summary>
    /// 获取当前词的含义数量
    /// </summary>
    int MeaningCount { get; }

    /// <summary>
    /// 获取当前词的所有含义列表
    /// </summary>
    object MeaningList { get; }

    /// <summary>
    /// 获取当前词的词性列表
    /// </summary>
    object PartOfSpeechList { get; }

    /**
        /// <summary>
        /// 获取当前词的同义词列表
        /// </summary>
        //object SynonymList { get; }
    **/

    /// <summary>
    /// 获取当前词的反义词列表
    /// </summary>
    object AntonymList { get; }

    /// <summary>
    /// 获取与当前词相关的表达列表
    /// </summary>
    object RelatedExpressionList { get; }

    /// <summary>
    /// 获取与当前词相关的词汇列表
    /// </summary>
    object RelatedWordList { get; }

    /// <summary>
    /// 获取当前处理的词
    /// </summary>
    string? Word { get; }
}
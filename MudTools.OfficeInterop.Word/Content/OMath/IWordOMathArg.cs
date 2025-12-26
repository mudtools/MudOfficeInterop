//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中数学对象参数(OMath Arguments)的集合接口
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord"), ItemIndex, NoneEnumerable]
public interface IWordOMathArgs : IEnumerable<IWordOMath?>, IOfficeObject<IWordOMathArgs>, IDisposable
{
    /// <summary>
    /// 获取与该数学对象参数集合关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此数学对象参数集合的父级对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中数学对象参数的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取集合中的指定数学对象参数
    /// </summary>
    /// <param name="index">要获取的数学对象参数的索引（从0开始）</param>
    /// <returns>指定索引处的数学对象参数，如果不存在则返回null</returns>
    IWordOMath? this[int index] { get; }

    /// <summary>
    /// 在当前数学对象参数集合中添加一个新的数学对象参数
    /// </summary>
    /// <param name="beforeArg">指定新参数应插入到的位置，如果为null则将新参数添加到集合末尾</param>
    /// <returns>新添加的数学对象参数，如果添加失败则返回null</returns>
    IWordOMath? Add(IWordOMathArgs? beforeArg = null);
}
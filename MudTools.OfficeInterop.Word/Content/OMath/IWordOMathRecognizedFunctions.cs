//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// OMathRecognizedFunctions 接口及实现类
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordOMathRecognizedFunctions : IEnumerable<IWordOMathRecognizedFunction?>, IOfficeObject<IWordOMathRecognizedFunctions>, IDisposable
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取集合中数学识别函数的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 返回集合中指定的 <see cref="IWordOMathRecognizedFunction"/> 对象。
    /// </summary>
    /// <param name="index">要返回的单个对象。可以是代表序号位置的 Number 类型的值。</param>
    /// <returns>指定索引处的 <see cref="IWordOMathRecognizedFunction"/> 对象。</returns>
    IWordOMathRecognizedFunction? this[int index] { get; }

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 将新的数学识别函数添加到集合中。
    /// </summary>
    /// <param name="Name">要添加的函数名称。</param>
    /// <returns>返回新添加的 <see cref="IWordOMathRecognizedFunction"/> 对象。</returns>
    IWordOMathRecognizedFunction? Add(string Name);
}
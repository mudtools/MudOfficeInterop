//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// PowerPoint 序列接口
/// </summary>
public interface IPowerPointSequences : IDisposable
{
    /// <summary>
    /// 获取序列数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 根据索引获取序列
    /// </summary>
    IPowerPointSequence this[int index] { get; }

    /// <summary>
    /// 添加序列
    /// </summary>
    /// <param name="index">插入位置</param>
    /// <returns>新添加的序列</returns>
    IPowerPointSequence Add(int index = -1);


    /// <summary>
    /// 根据条件查找序列
    /// </summary>
    /// <param name="predicate">查找条件</param>
    /// <returns>符合条件的序列列表</returns>
    IEnumerable<IPowerPointSequence> Find(Func<IPowerPointSequence, bool> predicate);

    /// <summary>
    /// 获取主序列
    /// </summary>
    /// <returns>主序列</returns>
    IPowerPointSequence GetMainSequence();

    /// <summary>
    /// 获取交互序列
    /// </summary>
    /// <returns>交互序列</returns>
    IEnumerable<IPowerPointSequence> GetInteractiveSequences();

    /// <summary>
    /// 重新排序序列
    /// </summary>
    /// <param name="newOrder">新顺序数组</param>
    void Reorder(int[] newOrder);
}
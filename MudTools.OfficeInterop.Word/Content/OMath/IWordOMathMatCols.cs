//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中数学矩阵的列集合
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordOMathMatCols : IEnumerable<IWordOMathMatCol?>, IDisposable
{
    /// <summary>
    /// 获取与当前对象关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }


    /// <summary>
    /// 获取当前对象的父级对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中矩阵列的数量
    /// </summary>
    int Count { get; }


    /// <summary>
    /// 根据索引获取指定的矩阵列
    /// </summary>
    /// <param name="index">要获取的矩阵列的索引</param>
    /// <returns>指定索引处的矩阵列，如果不存在则返回null</returns>
    IWordOMathMatCol? this[int index] { get; }

    /// <summary>
    /// 向集合中添加一个新的矩阵列
    /// </summary>
    /// <param name="beforeCol">要在其前面插入新列的现有列，如果为null则添加到末尾</param>
    /// <returns>新添加的矩阵列，如果操作失败则返回null</returns>
    IWordOMathMatCol? Add(IWordOMathMatCol? beforeCol);
}
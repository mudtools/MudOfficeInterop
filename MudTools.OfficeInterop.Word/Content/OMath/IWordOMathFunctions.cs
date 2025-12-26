//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的数学函数集合，提供对集合中数学函数的访问和管理功能
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordOMathFunctions : IEnumerable<IWordOMathFunction?>, IOfficeObject<IWordOMathFunctions>, IDisposable
{
    /// <summary>
    /// 获取与当前数学函数集合关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }


    /// <summary>
    /// 获取当前数学函数集合的父对象
    /// </summary>
    object? Parent { get; }


    /// <summary>
    /// 获取集合中数学函数的数量
    /// </summary>
    int Count { get; }


    /// <summary>
    /// 根据索引获取集合中的特定数学函数
    /// </summary>
    /// <param name="index">要获取的数学函数的索引，从0开始</param>
    /// <returns>指定索引处的数学函数对象，如果不存在则返回null</returns>
    IWordOMathFunction? this[int index] { get; }

    /// <summary>
    /// 向集合中添加一个新的数学函数
    /// </summary>
    /// <param name="range">用于创建数学函数的Word范围</param>
    /// <param name="type">数学函数的类型，使用WdOMathFunctionType枚举值</param>
    /// <param name="numArgs">函数参数的数量（可选）</param>
    /// <param name="numCols">列的数量（可选，主要用于矩阵等函数）</param>
    /// <returns>新创建的数学函数对象，如果创建失败则返回null</returns>
    IWordOMathFunction? Add(IWordRange range, WdOMathFunctionType type, int? numArgs = null, int? numCols = null);

}
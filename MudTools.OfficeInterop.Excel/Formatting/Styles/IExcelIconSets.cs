//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel中图标集集合的接口，提供对图标集集合的访问和操作功能
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelIconSets : IEnumerable<IExcelIconSet?>, IDisposable
{

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取集合中的图标集数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取指定位置的图标集
    /// </summary>
    /// <param name="index">图标集在集合中的索引位置</param>
    /// <returns>指定索引位置的图标集</returns>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelIconSet? this[int index] { get; }

    /// <summary>
    /// 通过名称获取指定的图标集
    /// </summary>
    /// <param name="name">图标集的名称</param>
    /// <returns>具有指定名称的图标集</returns>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelIconSet? this[string name] { get; }
}
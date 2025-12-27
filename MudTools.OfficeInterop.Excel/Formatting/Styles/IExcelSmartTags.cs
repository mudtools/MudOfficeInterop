//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel SmartTags 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.SmartTags 的安全访问和操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelSmartTags : IEnumerable<IExcelSmartTag?>, IDisposable
{
    /// <summary>
    /// 获取图例的父对象
    /// 对应 SmartTag.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取图例所在的 Application 对象
    /// 对应 SmartTag.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取智能标记集合中的智能标记数量
    /// 对应 SmartTags.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的智能标记对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">智能标记索引（从1开始）</param>
    /// <returns>智能标记对象</returns>
    IExcelSmartTag? this[int index] { get; }

    /// <summary>
    /// 向集合中添加新的智能标记
    /// </summary>
    /// <param name="smartTagType">智能标记类型</param>
    /// <returns>新创建的智能标记对象</returns>
    IExcelSmartTag? Add(string smartTagType);

}
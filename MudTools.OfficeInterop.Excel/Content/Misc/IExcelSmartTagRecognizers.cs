//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// SmartTagRecognizer 对象的集合，表示在 Microsoft Excel 中工作时将数据标记为信息类型的识别引擎。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelSmartTagRecognizers : IEnumerable<IExcelSmartTagRecognizer?>, IOfficeObject<IExcelSmartTagRecognizers, MsExcel.SmartTagRecognizers>, IDisposable
{
    /// <summary>
    /// 获取对象的父对象 
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取对象数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引处的 SmartTagRecognizer 对象。
    /// </summary>
    /// <param name="index">对象的名称或索引号。</param>
    /// <returns>指定索引处的 SmartTagRecognizer 对象。</returns>
    IExcelSmartTagRecognizer? this[int index] { get; }

    /// <summary>
    /// 获取指定索引处的 SmartTagRecognizer 对象。
    /// </summary>
    /// <param name="name">对象的名称或索引号。</param>
    /// <returns>指定索引处的 SmartTagRecognizer 对象。</returns>
    IExcelSmartTagRecognizer? this[string name] { get; }

    /// <summary>
    /// 获取或设置一个值，指示数据是否可以标记为智能标记。
    /// </summary>
    bool Recognize { get; set; }
}
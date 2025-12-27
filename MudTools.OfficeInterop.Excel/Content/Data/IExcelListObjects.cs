//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelListObjects : IOfficeObject<IExcelListObjects>, IDisposable, IEnumerable<IExcelListObject>
{
    /// <summary>
    /// 获取所属的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取列表对象的计数
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取列表对象
    /// </summary>
    /// <param name="index">索引</param>
    /// <returns>列表对象</returns>
    IExcelListObject? this[int index] { get; }

    /// <summary>
    /// 通过名称获取列表对象
    /// </summary>
    /// <param name="name">名称</param>
    /// <returns>列表对象</returns>
    IExcelListObject? this[string name] { get; }

    /// <summary>
    /// 添加新的列表对象
    /// </summary>
    /// <param name="sourceType">源类型</param>
    /// <param name="source">源数据</param>
    /// <param name="link">是否链接</param>
    /// <param name="xlListObjectHasHeaders">是否有标题行</param>
    /// <param name="destination">目标位置</param>
    /// <returns>新创建的列表对象</returns>
    IExcelListObject? AddEx(XlListObjectSourceType sourceType, object source, object link, XlYesNoGuess xlListObjectHasHeaders = XlYesNoGuess.xlGuess, object? destination = null);

}
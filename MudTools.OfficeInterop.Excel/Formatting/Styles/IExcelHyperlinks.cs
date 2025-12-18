//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// Excel Hyperlinks 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Hyperlinks 的安全访问和操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelHyperlinks : IEnumerable<IExcelHyperlink?>, IDisposable
{
    /// <summary>
    /// 获取超链接集合中的超链接数量
    /// 对应 Hyperlinks.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的超链接对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">超链接索引（从1开始）</param>
    /// <returns>超链接对象</returns>
    IExcelHyperlink? this[int index] { get; }

    /// <summary>
    /// 向集合中添加新的超链接
    /// </summary>
    /// <param name="anchor">超链接的锚点区域</param>
    /// <param name="address">链接地址（如网页URL或文件路径）</param>
    /// <param name="subAddress">子地址（如工作表名称或单元格引用）</param>
    /// <param name="screenTip">鼠标悬停时显示的提示文本</param>
    /// <param name="textToDisplay">要显示的文本</param>
    /// <returns>新创建的超链接对象</returns>
    [ReturnValueConvert]
    IExcelHyperlink? Add(IExcelRange anchor, string address, string? subAddress = null, string? screenTip = null, string? textToDisplay = null);

    /// <summary>
    /// 删除集合中的所有超链接
    /// 对应 Hyperlinks.Delete 方法
    /// </summary>
    void Delete();

}
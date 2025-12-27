//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel工作表页面的页眉和页脚设置接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPage : IOfficeObject<IExcelPage>, IDisposable
{
    /// <summary>
    /// 获取左页眉设置
    /// </summary>
    IExcelHeaderFooter? LeftHeader { get; }

    /// <summary>
    /// 获取中页眉设置
    /// </summary>
    IExcelHeaderFooter? CenterHeader { get; }

    /// <summary>
    /// 获取右页眉设置
    /// </summary>
    IExcelHeaderFooter? RightHeader { get; }

    /// <summary>
    /// 获取左页脚设置
    /// </summary>
    IExcelHeaderFooter? LeftFooter { get; }

    /// <summary>
    /// 获取中页脚设置
    /// </summary>
    IExcelHeaderFooter? CenterFooter { get; }

    /// <summary>
    /// 获取右页脚设置
    /// </summary>
    IExcelHeaderFooter? RightFooter { get; }

}
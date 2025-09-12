//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档或节中的行号设置。
/// 封装了 Microsoft.Office.Interop.Word.LineNumbering 对象。
/// </summary>
public interface IWordLineNumbering : IDisposable
{
    #region 属性

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建指定对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取代表 <see cref="IWordLineNumbering"/> 对象的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置行号的起始值。
    /// </summary>
    int? StartingNumber { get; set; }

    /// <summary>
    /// 获取或设置行号之间的距离（以行为单位）。
    /// </summary>
    int? CountBy { get; set; }


    /// <summary>
    /// 获取或设置行号的重启方式。
    /// </summary>
    WdNumberingRule RestartMode { get; set; }

    /// <summary>
    /// 获取或设置行号相对于页面或文本边界的水平位置（以磅为单位）。
    /// </summary>
    float? DistanceFromText { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，该值指示是否为每个页面或节重新开始编号。
    /// </summary>
    /// <remarks>
    /// 注意：此属性可能与 RestartMode 属性相关联或重叠。
    /// 具体行为取决于 Word 版本和上下文。建议优先使用 RestartMode。
    /// </remarks>
    int? Active { get; set; }
    #endregion 
}
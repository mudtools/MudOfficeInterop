//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 文档中的一个页面布局信息。
/// 封装了 Microsoft.Office.Interop.Word.Page 对象。
/// </summary>
/// <remarks>
/// Page 对象代表文档页面的布局信息，通常在打印或页面布局视图中生成。
/// 它提供对页面尺寸、边界和内容范围的访问。
/// </remarks>
public interface IWordPage : IDisposable
{
    #region 属性

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取页面的宽度（以磅为单位）。
    /// </summary>
    float? Width { get; }

    /// <summary>
    /// 获取页面的高度（以磅为单位）。
    /// </summary>
    float? Height { get; }

    /// <summary>
    /// 获取页面左边距（以磅为单位）。
    /// </summary>
    float? Left { get; }

    /// <summary>
    /// 获取页面上边距（以磅为单位）。
    /// </summary>
    float? Top { get; }

    /// <summary>
    /// 获取代表 <see cref="IWordPage"/> 对象的父对象（通常是 Pages 集合）。
    /// </summary>
    object? Parent { get; }
    #endregion

}
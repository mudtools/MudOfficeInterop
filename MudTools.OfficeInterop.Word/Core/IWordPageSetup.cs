//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;


/// <summary>
/// Word 文档页面设置接口
/// </summary>
public interface IWordPageSetup : IDisposable
{
    /// <summary>
    /// 获取或设置上边距
    /// </summary>
    float TopMargin { get; set; }

    /// <summary>
    /// 获取或设置下边距
    /// </summary>
    float BottomMargin { get; set; }

    /// <summary>
    /// 获取或设置左边距
    /// </summary>
    float LeftMargin { get; set; }

    /// <summary>
    /// 获取或设置右边距
    /// </summary>
    float RightMargin { get; set; }

    /// <summary>
    /// 获取或设置页面宽度
    /// </summary>
    float PageWidth { get; set; }

    /// <summary>
    /// 获取或设置页面高度
    /// </summary>
    float PageHeight { get; set; }

    /// <summary>
    /// 获取或设置页面方向（0=纵向，1=横向）
    /// </summary>
    int Orientation { get; set; }
}
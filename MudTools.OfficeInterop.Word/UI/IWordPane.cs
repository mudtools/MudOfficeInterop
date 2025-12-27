//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档窗口中的一个窗格。
/// 封装了 Microsoft.Office.Interop.Word.Pane 对象。
/// </summary>
public interface IWordPane : IDisposable
{
    #region 属性

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取或设置窗格中文本的垂直滚动位置（以磅为单位）。
    /// </summary>
    int? VerticalPercentScrolled { get; set; }

    /// <summary>
    /// 获取窗格的索引号。
    /// </summary>
    int? Index { get; }

    /// <summary>
    /// 获取窗格中文本的水平滚动位置（以磅为单位）。
    /// </summary>
    int? HorizontalPercentScrolled { get; set; }

    /// <summary>
    /// 获取窗格的文档。
    /// </summary>
    IWordDocument? Document { get; }

    /// <summary>
    /// 获取窗格的文档视图设置。
    /// </summary>
    IWordView? View { get; }

    /// <summary>
    /// 获取代表 <see cref="IWordPane"/> 对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取窗格的选定内容。
    /// </summary>
    IWordSelection? Selection { get; }

    IWordPages? Pages { get; }

    IWordZooms? Zooms { get; }

    #endregion // 属性

    #region 方法

    /// <summary>
    /// 激活指定的窗格。
    /// </summary>
    void Activate();

    void AutoScroll(int velocity);

    void Close();

    void NewFrameset();

    /// <summary>
    /// 将窗格滚动到下一页。
    /// </summary>
    void LargeScroll(int? down = null, int? up = null, int? toRight = null, int? toLeft = null);

    /// <summary>
    /// 将窗格滚动到下一页。
    /// </summary>
    void PageScroll(int? pages = null, int? lines = null);

    /// <summary>
    /// 将窗格滚动到下一页。
    /// </summary>
    void SmallScroll(int? down = null, int? up = null, int? toRight = null, int? toLeft = null);

    #endregion // 方法
}
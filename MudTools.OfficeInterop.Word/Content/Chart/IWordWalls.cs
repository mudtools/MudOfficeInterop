//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表墙壁的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordWalls : IOfficeObject<IWordWalls>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置墙壁名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置厚度。
    /// </summary>
    int Thickness { get; set; }

    /// <summary>
    /// 获取或设置图片单位大小。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    int PictureUnit { get; set; }

    /// <summary>
    /// 获取内部区域格式。
    /// </summary>
    IWordInterior? Interior { get; }

    /// <summary>
    /// 获取填充格式。
    /// </summary>
    IWordChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取边框格式。
    /// </summary>
    IWordChartBorder? Border { get; }

    /// <summary>
    /// 获取格式对象。
    /// </summary>
    IWordChartFormat? Format { get; }

    /// <summary>
    /// 获取或设置是否显示图片背面。
    /// </summary>
    XlChartPictureType PictureType { get; set; }

    /// <summary>
    /// 选择墙壁。
    /// </summary>
    void Select();

    /// <summary>
    /// 将剪贴板的内容粘贴到墙壁上。
    /// </summary>
    void Paste();

    /// <summary>
    /// 清除指定对象的格式设置并将格式设置恢复为默认值。
    /// </summary>
    /// <returns>返回清除格式后的对象。</returns>
    object? ClearFormats();
}
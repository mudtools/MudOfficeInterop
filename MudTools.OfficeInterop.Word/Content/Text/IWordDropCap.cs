//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示段落开头的首字下沉字母。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordDropCap : IOfficeObject<IWordDropCap, MsWord.DropCap>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置首字下沉字母的位置。
    /// </summary>
    WdDropPosition Position { get; set; }

    /// <summary>
    /// 获取或设置首字下沉字母的字体名称。
    /// </summary>
    string FontName { get; set; }

    /// <summary>
    /// 获取或设置指定首字下沉字母的高度（以行为单位）。
    /// </summary>
    int LinesToDrop { get; set; }

    /// <summary>
    /// 获取或设置首字下沉字母与段落文本之间的距离（以磅为单位）。
    /// </summary>
    float DistanceFromText { get; set; }

    /// <summary>
    /// 移除首字下沉字母的格式设置。
    /// </summary>
    void Clear();

    /// <summary>
    /// 将指定段落的第一个字符格式化为首字下沉字母。
    /// </summary>
    void Enable();
}
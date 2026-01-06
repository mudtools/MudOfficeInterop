//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 文档中的一个修订（Revision）的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordRevision : IOfficeObject<IWordRevision>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取修订的作者。
    /// </summary>
    string Author { get; }

    /// <summary>
    /// 获取修订的日期和时间。
    /// </summary>
    DateTime Date { get; }

    /// <summary>
    /// 获取修订的类型。
    /// </summary>
    WdRevisionType Type { get; }

    /// <summary>
    /// 获取修订的范围。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取或设置修订的格式描述。
    /// </summary>
    string FormatDescription { get; }

    /// <summary>
    /// 获取修订的索引号。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取修订的样式。
    /// </summary>
    IWordStyle? Style { get; }

    /// <summary>
    /// 获取移动修订的原始范围。
    /// </summary>
    IWordRange? MovedRange { get; }

    /// <summary>
    /// 获取修改的表格。
    /// </summary>
    IWordCells? Cells { get; }

    /// <summary>
    /// 接受此修订。
    /// </summary>
    void Accept();

    /// <summary>
    /// 拒绝此修订。
    /// </summary>
    void Reject();
}
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中的一个邮件合并域（如 MERGEFIELD 域）的二次封装接口。
/// 此接口提供了对该域的范围、数据源字段名称以及操作的访问 [[1]]。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordMailMergeField : IDisposable
{
    /// <summary>
    /// 获取此邮件合并域所属的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此邮件合并域的父对象（通常是 <see cref="IWordMailMergeFields"/> 集合或 <see cref="IWordRange"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此邮件合并域在文档中所占据的范围 (<see cref="IWordRange"/>)。
    /// </summary>
    IWordRange? Code { get; }

    /// <summary>
    /// 获取文档中的下一个邮件合并域。
    /// </summary>
    IWordMailMergeField? Next { get; }

    /// <summary>
    /// 获取文档中的上一个邮件合并域。
    /// </summary>
    IWordMailMergeField? Previous { get; }

    /// <summary>
    /// 获取域的类型。
    /// </summary>
    WdFieldType Type { get; }

    /// <summary>
    /// 获取或设置域是否被锁定。锁定的域在文档更新时不会改变。
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 从文档中删除此邮件合并域。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选中此邮件合并域。
    /// </summary>
    void Select();

    /// <summary>
    /// 剪切此邮件合并域。
    /// </summary>
    void Cut();

    /// <summary>
    /// 复制此邮件合并域。
    /// </summary>
    void Copy();
}
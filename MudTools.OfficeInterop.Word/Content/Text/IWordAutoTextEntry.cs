//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 中的单个“自动图文集”条目（AutoText Entry）的封装接口。
/// 自动图文集条目允许用户通过简短名称快速插入预定义内容（如段落、表格等）。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordAutoTextEntry : IDisposable
{
    /// <summary>
    /// 获取此对象所属的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此对象的父对象（通常是 <see cref="IWordMailMerge"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取自动图文集条目的名称（如 "Address"）
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取自动图文集条目在集合中的索引位置（从1开始计数）
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置自动图文集条目的值（即插入的内容）
    /// </summary>
    string Value { get; set; }

    /// <summary>
    ///获取此自动图文集条目样式名
    /// </summary>
    string StyleName { get; }


    /// <summary>
    /// 从模板中删除此自动图文集条目
    /// </summary>
    void Delete();

    /// <summary>
    /// 在指定位置插入自动图文集条目内容
    /// </summary>
    /// <param name="where">要插入内容的位置</param>
    /// <param name="richText">如果为 True，则将构建基块作为 RTF 格式文本插入。 如果为 False，则将构建基块作为纯文本插入。</param>
    /// <returns>插入内容后的新范围</returns>
    IWordRange? Insert(IWordRange where, bool? richText = null);

}
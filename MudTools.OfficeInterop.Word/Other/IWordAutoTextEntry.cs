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
public interface IWordAutoTextEntry : IDisposable
{
    /// <summary>
    /// 获取自动图文集条目的名称（如 "Address"）
    /// </summary>
    string Name { get; }

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
}
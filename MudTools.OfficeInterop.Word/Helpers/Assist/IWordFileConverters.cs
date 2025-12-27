//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示可用于打开和保存文件的所有文件转换器的集合。
/// <para>注：使用 Item[Object] 方法或 FileConverters(index) 可返回单个 FileConverter 对象。
/// 索引号代表文件转换器在 FileConverters 集合中的位置。</para>
/// <para>注：Add 方法不适用于 FileConverters 集合。FileConverter 对象是在安装 Microsoft Office 或附加文件转换器时添加的。</para>
/// </summary>
public interface IWordFileConverters : IEnumerable<IWordFileConverter>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中的文件转换器数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引（类名或索引号）获取单个文件转换器。
    /// </summary>
    /// <param name="index">类名（字符串）或索引号（整数）。</param>
    /// <returns>指定的文件转换器对象。</returns>
    IWordFileConverter this[object index] { get; }

    /// <summary>
    /// 获取或设置一个值，该值控制是否将 Microsoft Word 6.0/95/97 中的 V 型字符 (??) 转换为 Microsoft Word 中的引号标记。
    /// </summary>
    WdChevronConvertRule ConvertMacWordChevrons { get; set; }
}
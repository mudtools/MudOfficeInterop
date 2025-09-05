//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示用于打开或保存文件的文件转换器。
/// <para>注：FileConverter 对象是 FileConverters 集合的成员。</para>
/// <para>注：不能创建新的文件转换器，也不能向 FileConverters 集合中添加新的文件转换器。
/// FileConverter 对象是在安装 Microsoft Office 或附加文件转换器时添加的。</para>
/// </summary>
public interface IWordFileConverter : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取一个值，该值指示指定的文件转换器是否设计为打开文件。
    /// </summary>
    bool CanOpen { get; }

    /// <summary>
    /// 获取一个值，该值指示指定的文件转换器是否设计为保存文件。
    /// </summary>
    bool CanSave { get; }

    /// <summary>
    /// 获取唯一标识文件转换器的类名。
    /// </summary>
    string ClassName { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取与指定 FileConverter 对象关联的文件扩展名。
    /// </summary>
    string Extensions { get; }

    /// <summary>
    /// 获取指定的文件转换器的显示名称。
    /// </summary>
    string FormatName { get; }

    /// <summary>
    /// 获取或设置指定对象的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取指定的文件转换器用于打开文件的文件格式。
    /// </summary>
    int OpenFormat { get; }

    /// <summary>
    /// 获取指定的文件转换器用于保存文件的文件格式。
    /// </summary>
    int SaveFormat { get; }

    /// <summary>
    /// 获取指定对象的磁盘或 Web 路径。
    /// </summary>
    string Path { get; }
}
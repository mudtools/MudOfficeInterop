//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 中 OLE 对象格式的封装接口
/// </summary>
public interface IWordOLEFormat : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取 OLE 对象的类名
    /// </summary>
    string ClassType { get; }

    /// <summary>
    /// 获取 OLE 对象的程序标识符
    /// </summary>
    string ProgID { get; }

    /// <summary>
    /// 获取或设置 OLE 对象的图标索引
    /// </summary>
    int IconIndex { get; set; }

    /// <summary>
    /// 获取或设置 OLE 对象的图标标签
    /// </summary>
    string IconLabel { get; set; }

    /// <summary>
    /// 获取 OLE 对象是否为链接对象
    /// </summary>
    bool IsLinked { get; }

    /// <summary>
    /// 获取 OLE 对象是否为嵌入对象
    /// </summary>
    bool IsEmbedded { get; }

    /// <summary>
    /// 获取或设置 OLE 对象是否以图标形式显示
    /// </summary>
    bool DisplayAsIcon { get; set; }

    /// <summary>
    /// 获取 OLE 对象的原始格式（伪代码）
    /// </summary>
    object Object { get; }

    /// <summary>
    /// 获取 OLE 对象的应用程序对象（伪代码）
    /// </summary>
    object Application { get; }

    /// <summary>
    /// 激活 OLE 对象以进行编辑
    /// </summary>
    void Activate();

    /// <summary>
    /// 编辑 OLE 对象
    /// </summary>
    /// <param name="verb">要执行的动作动词索引</param>
    void DoVerb(int verb = 1);
}
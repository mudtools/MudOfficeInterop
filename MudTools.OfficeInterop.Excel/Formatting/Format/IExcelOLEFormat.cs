//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 中 OLE 对象格式的封装接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelOLEFormat : IOfficeObject<IExcelOLEFormat>, IDisposable
{

    /// <summary>
    /// 获取 OLE 对象的程序标识符
    /// </summary>
    string progID { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 返回与此 OLE 对象相联系的 OLE 自动化对象。
    /// </summary>
    object Object { get; }

    /// <summary>
    /// 激活 OLE 对象以进行编辑
    /// </summary>
    void Activate();

    /// <summary>
    /// 向指定的 OLE 对象服务器发送动词。
    /// </summary>
    /// <param name="verb">可选 对象。 OLE 对象服务器将执行其操作的动词。 如果省略此参数，则发送默认动词。</param>
    void Verb(XlOLEVerb? verb = null);
}
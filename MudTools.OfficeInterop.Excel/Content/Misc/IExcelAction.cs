//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示要在数据透视表或工作表数据中执行的操作。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelAction : IOfficeObject<IExcelAction, MsExcel.Action>, IDisposable
{

    /// <summary>
    /// 获取当前COM对象的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取当前COM对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取分配给 Action 对象的标题。只读。
    /// </summary>
    string Caption { get; }

    /// <summary>
    /// 获取操作类型。只读。
    /// </summary>
    XlActionType Type { get; }

    /// <summary>
    /// 获取 Action 对象的坐标属性。只读。
    /// </summary>
    string Coordinate { get; }

    /// <summary>
    /// 获取与 Action 对象关联的内容。只读。
    /// </summary>
    string Content { get; }

    /// <summary>
    /// 执行指定的操作。
    /// </summary>
    void Execute();
}
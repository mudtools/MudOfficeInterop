//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel SmartTag 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.SmartTag 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelSmartTag : IOfficeObject<IExcelSmartTag>, IDisposable
{
    /// <summary>
    /// 获取图例的父对象
    /// 对应 SmartTag.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取图例所在的 Application 对象
    /// 对应 SmartTag.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取智能标记的名称
    /// 对应 SmartTag.Name 属性
    /// </summary>
    string? Name { get; }

    /// <summary>
    /// 获取智能标记的XML字符串
    /// 对应 SmartTag.XML 属性
    /// </summary>
    string? XML { get; }

    /// <summary>
    /// 获取智能标记所在的区域对象
    /// 对应 SmartTag.Range 属性
    /// </summary>
    IExcelRange? Range { get; }

    /// <summary>
    /// 获取智能标记的动作集合
    /// 对应 SmartTag.SmartTagActions 属性
    /// </summary>
    IExcelSmartTagActions? SmartTagActions { get; }

    /// <summary>
    /// 获取智能标记的自定义属性集合
    /// 对应 SmartTag.Properties 属性
    /// </summary>
    IExcelCustomProperties? Properties { get; }

    /// <summary>
    /// 删除智能标记
    /// 对应 SmartTag.Delete 方法
    /// </summary>
    void Delete();
}
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel PivotItem 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PivotItem 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPivotItem : IOfficeObject<IExcelPivotItem>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取数据透视表项目的名称
    /// 对应 PivotItem.Name 属性
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取数据透视表项目的父对象 (通常是 PivotField)
    /// 对应 PivotItem.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取数据透视表项目所在的Application对象
    /// 对应 PivotItem.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置数据透视表项目是否可见
    /// 对应 PivotItem.Visible 属性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置数据透视表项目的公式
    /// 对应 PivotItem.Formula 属性
    /// </summary>
    string Formula { get; set; }

    /// <summary>
    /// 获取数据透视表项目的源名称
    /// 对应 PivotItem.SourceName 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string SourceName { get; }
    #endregion


    #region 图表元素 (子对象)
    /// <summary>
    /// 获取数据透视表项目的数据范围 (如果适用)
    /// 对应 PivotItem.DataRange 属性
    /// </summary>
    IExcelRange DataRange { get; }

    /// <summary>
    /// 获取数据透视表项目的标签范围 (如果适用)
    /// 对应 PivotItem.LabelRange 属性
    /// </summary>
    IExcelRange LabelRange { get; }
    #endregion

    /// <summary>
    /// 删除数据透视表项目
    /// 对应 PivotItem.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 钻取到指定的数据透视表字段
    /// 对应 PivotItem.DrillTo 方法
    /// </summary>
    /// <param name="field">要钻取到的字段名称</param>
    void DrillTo(string field);
}

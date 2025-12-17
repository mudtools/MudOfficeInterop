//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel TableStyle 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.TableStyle 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelTableStyle : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取表格样式的名称
    /// 对应 TableStyle.Name 属性
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取表格样式的本地化名称
    /// 对应 TableStyle.NameLocal 属性
    /// </summary>
    string NameLocal { get; }

    /// <summary>
    /// 获取表格样式的父对象 (通常是 Workbook)
    /// 对应 TableStyle.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取表格样式所在的Application对象
    /// 对应 TableStyle.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取表格样式是否为内置样式
    /// 对应 TableStyle.BuiltIn 属性
    /// </summary>
    bool IsBuiltIn { get; }
    #endregion

    /// <summary>
    /// 获取表格样式元素集合
    /// 对应 TableStyle.TableStyleElements 属性
    /// </summary>
    IExcelTableStyleElements? TableStyleElements { get; }

    /// <summary>
    /// 复制当前表格样式并创建一个新的表格样式
    /// 对应 TableStyle.Duplicate 方法
    /// </summary>
    /// <param name="newTableStyleName">新表格样式的名称</param>
    /// <returns>复制后的新表格样式对象</returns>
    IExcelTableStyle Duplicate(string? newTableStyleName);

    /// <summary>
    /// 获取或设置是否在表格样式库中显示该样式
    /// 对应 TableStyle.ShowAsAvailableTableStyle 属性
    /// </summary>
    bool ShowAsAvailableTableStyle { get; set; }

    /// <summary>
    /// 获取或设置是否在数据透视表样式库中显示该样式
    /// 对应 TableStyle.ShowAsAvailablePivotTableStyle 属性
    /// </summary>
    bool ShowAsAvailablePivotTableStyle { get; set; }

    /// <summary>
    /// 获取或设置是否在切片器样式库中显示该样式
    /// 对应 TableStyle.ShowAsAvailableSlicerStyle 属性
    /// </summary>
    bool ShowAsAvailableSlicerStyle { get; set; }

    /// <summary>
    /// 获取或设置是否在时间线样式库中显示该样式
    /// 对应 TableStyle.ShowAsAvailableTimelineStyle 属性
    /// </summary>
    bool ShowAsAvailableTimelineStyle { get; set; }

    /// <summary>
    /// 删除当前表格样式
    /// 对应 TableStyle.Delete 方法
    /// </summary>
    void Delete();
}

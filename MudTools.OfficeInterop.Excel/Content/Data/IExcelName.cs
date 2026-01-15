//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Name 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Name 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelName : IOfficeObject<IExcelName, MsExcel.Name>, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取名称所在的父对象
    /// 对应 Name.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取名称所在的Application对象
    /// 对应 Name.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置名称的值
    /// 对应 Name.Value 属性或 Name.RefersTo 属性
    /// </summary>
    string Value { get; set; }

    /// <summary>
    /// 获取或设置名称
    /// 对应 Name.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置本地名称
    /// 对应 Name.NameLocal 属性
    /// </summary>
    string NameLocal { get; set; }

    /// <summary>
    /// 获取名称的索引位置
    /// 对应 Name.Index 属性
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置引用
    /// 对应 Name.RefersTo 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string RefersTo { get; set; }

    /// <summary>
    /// 获取或设置本地引用
    /// 对应 Name.RefersToLocal 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string RefersToLocal { get; set; }

    /// <summary>
    /// 获取或设置R1C1引用
    /// 对应 Name.RefersToR1C1 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string RefersToR1C1 { get; set; }

    /// <summary>
    /// 获取或设置本地R1C1引用
    /// 对应 Name.RefersToR1C1Local 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string RefersToR1C1Local { get; set; }

    /// <summary>
    /// 获取或设置是否可见
    /// 对应 Name.Visible 属性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置类别
    /// 对应 Name.Category 属性
    /// </summary>
    string? Category { get; set; }

    /// <summary>
    /// 获取或设置本地类别
    /// 对应 Name.CategoryLocal 属性
    /// </summary>
    string? CategoryLocal { get; set; }

    /// <summary>
    /// 获取或设置宏类型
    /// 对应 Name.MacroType 属性
    /// </summary>
    XlXLMMacroType? MacroType { get; set; }

    /// <summary>
    /// 获取或设置快捷键
    /// 对应 Name.ShortcutKey 属性
    /// </summary>
    string ShortcutKey { get; set; }


    /// <summary>
    /// 获取引用的区域对象
    /// 对应 Name.RefersToRange 属性
    /// </summary>
    IExcelRange RefersToRange { get; }

    /// <summary>
    /// 获取或设置注释
    /// 对应 Name.Comment 属性
    /// </summary>
    string Comment { get; set; }

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除名称
    /// 对应 Name.Delete 方法
    /// </summary>
    void Delete();

    #endregion

}
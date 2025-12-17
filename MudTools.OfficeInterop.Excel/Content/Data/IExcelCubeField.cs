//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示 Excel 中的数据透视表 OLAP Cube 字段的接口
/// 该接口提供了对 Cube 字段的各种属性和操作方法的访问
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelCubeField : IDisposable
{
    /// <summary>
    /// 获取所在的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取 Cube 字段的类型
    /// </summary>
    XlCubeFieldType CubeFieldType { get; }

    /// <summary>
    /// 获取 Cube 字段的名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取 Cube 字段的值
    /// </summary>
    string Value { get; }

    /// <summary>
    /// 获取或设置字段在数据透视表中的方向（行、列、页等）
    /// </summary>
    XlPivotFieldOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置字段在数据透视表中的位置
    /// </summary>
    int Position { get; set; }

    /// <summary>
    /// 获取字段的树状视图控件
    /// </summary>
    IExcelTreeviewControl? TreeviewControl { get; }

    /// <summary>
    /// 获取或设置是否允许将字段拖拽到列区域
    /// </summary>
    bool DragToColumn { get; set; }

    /// <summary>
    /// 获取或设置是否允许将字段拖拽到隐藏区域
    /// </summary>
    bool DragToHide { get; set; }

    /// <summary>
    /// 获取或设置是否允许将字段拖拽到页面区域
    /// </summary>
    bool DragToPage { get; set; }

    /// <summary>
    /// 获取或设置是否允许将字段拖拽到行区域
    /// </summary>
    bool DragToRow { get; set; }

    /// <summary>
    /// 获取或设置是否允许将字段拖拽到数据区域
    /// </summary>
    bool DragToData { get; set; }

    /// <summary>
    /// 获取或设置隐藏级别数
    /// </summary>
    int HiddenLevels { get; set; }

    /// <summary>
    /// 获取字段是否有成员属性
    /// </summary>
    bool HasMemberProperties { get; }

    /// <summary>
    /// 获取或设置字段的布局形式
    /// </summary>
    XlLayoutFormType LayoutForm { get; set; }

    /// <summary>
    /// 获取字段关联的数据透视表字段集合
    /// </summary>
    IExcelPivotFields? PivotFields { get; }

    /// <summary>
    /// 获取或设置是否允许多个页面项
    /// </summary>
    bool EnableMultiplePageItems { get; set; }

    /// <summary>
    /// 获取或设置字段是否显示在字段列表中
    /// </summary>
    bool ShowInFieldList { get; set; }

    /// <summary>
    /// 获取或设置布局中小计的位置
    /// </summary>
    XlSubtototalLocationType LayoutSubtotalLocation { get; set; }

    /// <summary>
    /// 获取或设置 Cube 字段的子类型
    /// </summary>
    XlCubeFieldSubType CubeFieldSubType { get; }

    /// <summary>
    /// 获取或设置是否在筛选器中包含新项
    /// </summary>
    bool IncludeNewItemsInFilter { get; set; }

    /// <summary>
    /// 获取所有项是否可见
    /// </summary>
    bool AllItemsVisible { get; }

    /// <summary>
    /// 获取或设置当前页面的名称
    /// </summary>
    string CurrentPageName { get; set; }

    /// <summary>
    /// 获取字段是否表示日期
    /// </summary>
    bool IsDate { get; }

    /// <summary>
    /// 获取或设置字段的标题
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// 获取或设置是否展平层次结构
    /// </summary>
    bool FlattenHierarchies { get; set; }

    /// <summary>
    /// 获取或设置是否区分层次结构
    /// </summary>
    bool HierarchizeDistinct { get; set; }

    /// <summary>
    /// 添加成员属性字段
    /// </summary>
    /// <param name="property">要添加的属性</param>
    /// <param name="propertyOrder">属性的顺序</param>
    void AddMemberPropertyField(string property, int? propertyOrder = null);

    /// <summary>
    /// 创建数据透视表字段
    /// </summary>
    void CreatePivotFields();

    /// <summary>
    /// 删除字段
    /// </summary>
    void Delete();

    /// <summary>
    /// 清除手动筛选
    /// </summary>
    void ClearManualFilter();

}
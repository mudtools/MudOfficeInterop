//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel PivotField 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PivotField 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPivotField : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取或设置数据透视表字段的名称
    /// 对应 PivotField.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取数据透视表字段的父对象 (通常是 PivotTable)
    /// 对应 PivotField.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取数据透视表字段所在的Application对象
    /// 对应 PivotField.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置数据透视表字段的方向 (行、列、页、数据或隐藏)
    /// 对应 PivotField.Orientation 属性
    /// </summary>
    XlPivotFieldOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置数据透视表字段的位置 (在其方向内的顺序)
    /// 对应 PivotField.Position 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    int Position { get; set; }

    /// <summary>
    /// 获取或设置数据透视表字段的数字格式
    /// 对应 PivotField.NumberFormat 属性
    /// </summary>
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取数据透视表字段的数据源范围
    /// 对应 PivotField.SourceName 属性 (可能需要解析)
    /// </summary>
    object SourceName { get; }

    /// <summary>
    /// 获取数据透视表字段的汇总函数
    /// 对应 PivotField.Function 属性
    /// </summary>
    XlConsolidationFunction Function { get; set; }
    #endregion

    #region 数据透视表字段属性

    /// <summary>
    /// 获取与数据透视表字段关联的多维数据集字段
    /// 对应 PivotField.CubeField 属性
    /// </summary>
    IExcelCubeField CubeField { get; }

    /// <summary>
    /// 获取此字段的父属性字段
    /// 对应 PivotField.PropertyParentField 属性
    /// </summary>
    IExcelPivotField PropertyParentField { get; }

    /// <summary>
    /// 获取数据透视表字段的筛选器集合
    /// 对应 PivotField.PivotFilters 属性
    /// </summary>
    IExcelPivotFilters PivotFilters { get; }

    /// <summary>
    /// 获取自动排序所依据的数据透视表线
    /// 对应 PivotField.AutoSortPivotLine 属性
    /// </summary>
    IExcelPivotLine AutoSortPivotLine { get; }

    /// <summary>
    /// 获取当前字段的子字段
    /// 对应 PivotField.ChildField 属性
    /// </summary>
    IExcelPivotField ChildField { get; }

    /// <summary>
    /// 获取当前字段的父字段
    /// 对应 PivotField.ParentField 属性
    /// </summary>
    IExcelPivotField ParentField { get; }

    /// <summary>
    /// 获取或设置是否显示所有项（包括未在筛选器中选择的项）
    /// 对应 PivotField.ShowAllItems 属性
    /// </summary>
    bool ShowAllItems { get; set; }

    /// <summary>
    /// 获取或设置字段的公式
    /// 对应 PivotField.Formula 属性
    /// </summary>
    string Formula { get; set; }

    /// <summary>
    /// 获取或设置属性字段的顺序
    /// 对应 PivotField.PropertyOrder 属性
    /// </summary>
    int PropertyOrder { get; set; }

    /// <summary>
    /// 获取或设置当前页面列表
    /// 对应 PivotField.CurrentPageList 属性
    /// </summary>
    object CurrentPageList { get; set; }

    /// <summary>
    /// 获取字段是否正在使用内存
    /// 对应 PivotField.MemoryUsed 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool MemoryUsed { get; }

    /// <summary>
    /// 获取或设置字段是否基于服务器进行计算
    /// 对应 PivotField.ServerBased 属性
    /// </summary>
    bool ServerBased { get; set; }

    /// <summary>
    /// 获取或设置字段是否可以拖拽到列区域
    /// 对应 PivotField.DragToColumn 属性
    /// </summary>
    bool DragToColumn { get; set; }

    /// <summary>
    /// 获取或设置字段是否可以拖拽到隐藏区域
    /// 对应 PivotField.DragToHide 属性
    /// </summary>
    bool DragToHide { get; set; }

    /// <summary>
    /// 获取或设置字段是否可以拖拽到页面区域
    /// 对应 PivotField.DragToPage 属性
    /// </summary>
    bool DragToPage { get; set; }

    /// <summary>
    /// 获取或设置字段是否可以拖拽到行区域
    /// 对应 PivotField.DragToRow 属性
    /// </summary>
    bool DragToRow { get; set; }

    /// <summary>
    /// 获取或设置字段是否可以拖拽到数据区域
    /// 对应 PivotField.DragToData 属性
    /// </summary>
    bool DragToData { get; set; }

    /// <summary>
    /// 获取或设置布局中是否显示空行
    /// 对应 PivotField.LayoutBlankLine 属性
    /// </summary>
    bool LayoutBlankLine { get; set; }

    /// <summary>
    /// 获取或设置布局中小计位置的类型
    /// 对应 PivotField.LayoutSubtotalLocation 属性
    /// </summary>
    XlSubtototalLocationType LayoutSubtotalLocation { get; set; }

    /// <summary>
    /// 获取或设置字段的计算方式
    /// 对应 PivotField.Calculation 属性
    /// </summary>
    XlPivotFieldCalculation Calculation { get; set; }

    /// <summary>
    /// 获取或设置布局形式类型
    /// 对应 PivotField.LayoutForm 属性
    /// </summary>
    XlLayoutFormType LayoutForm { get; set; }

    /// <summary>
    /// 获取字段的数据类型
    /// 对应 PivotField.DataType 属性
    /// </summary>
    XlPivotFieldDataType DataType { get; }

    /// <summary>
    /// 获取字段的分组级别
    /// 对应 PivotField.GroupLevel 属性
    /// </summary>
    object GroupLevel { get; }

    /// <summary>
    /// 获取或设置字段是否已展开钻取
    /// 对应 PivotField.DrilledDown 属性
    /// </summary>
    bool DrilledDown { get; set; }

    /// <summary>
    /// 获取或设置当前页面的名称
    /// 对应 PivotField.CurrentPageName 属性
    /// </summary>
    string CurrentPageName { get; set; }

    /// <summary>
    /// 获取或设置标准公式
    /// 对应 PivotField.StandardFormula 属性
    /// </summary>
    string StandardFormula { get; set; }

    /// <summary>
    /// 获取或设置隐藏项列表
    /// 对应 PivotField.HiddenItemsList 属性
    /// </summary>
    object HiddenItemsList { get; set; }

    /// <summary>
    /// 获取隐藏项集合
    /// 对应 PivotField.HiddenItems 属性
    /// </summary>
    object HiddenItems { get; }

    /// <summary>
    /// 获取父项集合
    /// 对应 PivotField.ParentItems 属性
    /// </summary>
    object ParentItems { get; }

    /// <summary>
    /// 获取或设置可见项列表
    /// 对应 PivotField.VisibleItemsList 属性
    /// </summary>
    object VisibleItemsList { get; set; }

    /// <summary>
    /// 获取子项集合
    /// 对应 PivotField.ChildItems 属性
    /// </summary>
    object ChildItems { get; }

    /// <summary>
    /// 获取或设置是否使用成员属性作为标题
    /// 对应 PivotField.UseMemberPropertyAsCaption 属性
    /// </summary>
    bool UseMemberPropertyAsCaption { get; set; }

    /// <summary>
    /// 获取或设置成员属性标题
    /// 对应 PivotField.MemberPropertyCaption 属性
    /// </summary>
    string MemberPropertyCaption { get; set; }

    /// <summary>
    /// 获取或设置当前页面
    /// 对应 PivotField.CurrentPage 属性
    /// </summary>
    object CurrentPage { get; set; }

    /// <summary>
    /// 获取或设置是否以工具提示形式显示
    /// 对应 PivotField.DisplayAsTooltip 属性
    /// </summary>
    bool DisplayAsTooltip { get; set; }

    /// <summary>
    /// 获取或设置布局是否紧凑排列行
    /// 对应 PivotField.LayoutCompactRow 属性
    /// </summary>
    bool LayoutCompactRow { get; set; }

    /// <summary>
    /// 获取或设置是否在筛选器中包含新项
    /// 对应 PivotField.IncludeNewItemsInFilter 属性
    /// </summary>
    bool IncludeNewItemsInFilter { get; set; }

    /// <summary>
    /// 获取或设置是否在报表中显示
    /// 对应 PivotField.DisplayInReport 属性
    /// </summary>
    bool DisplayInReport { get; set; }

    /// <summary>
    /// 获取字段是否显示为标题
    /// 对应 PivotField.DisplayAsCaption 属性
    /// </summary>
    bool DisplayAsCaption { get; }

    /// <summary>
    /// 获取或设置字段是否隐藏
    /// 对应 PivotField.Hidden 属性
    /// </summary>
    bool Hidden { get; set; }

    /// <summary>
    /// 获取或设置是否按数据库排序
    /// 对应 PivotField.DatabaseSort 属性
    /// </summary>
    bool DatabaseSort { get; set; }

    /// <summary>
    /// 获取或设置字段标题
    /// 对应 PivotField.Caption 属性
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// 获取或设置小计名称
    /// 对应 PivotField.SubtotalName 属性
    /// </summary>
    string SubtotalName { get; set; }

    /// <summary>
    /// 获取或设置布局是否有分页符
    /// 对应 PivotField.LayoutPageBreak 属性
    /// </summary>
    bool LayoutPageBreak { get; set; }

    /// <summary>
    /// 获取或设置是否启用多个页面项
    /// 对应 PivotField.EnableMultiplePageItems 属性
    /// </summary>
    bool EnableMultiplePageItems { get; set; }

    /// <summary>
    /// 获取或设置是否显示明细数据
    /// 对应 PivotField.ShowDetail 属性
    /// </summary>
    bool ShowDetail { get; set; }

    /// <summary>
    /// 获取或设置是否重复显示标签
    /// 对应 PivotField.RepeatLabels 属性
    /// </summary>
    bool RepeatLabels { get; set; }

    /// <summary>
    /// 获取自动排序顺序
    /// 对应 PivotField.AutoSortOrder 属性
    /// </summary>
    int AutoSortOrder { get; }

    /// <summary>
    /// 获取自动排序自定义小计
    /// 对应 PivotField.AutoSortCustomSubtotal 属性
    /// </summary>
    int AutoSortCustomSubtotal { get; }

    /// <summary>
    /// 获取字段是否在轴上显示
    /// 对应 PivotField.ShowingInAxis 属性
    /// </summary>
    bool ShowingInAxis { get; }

    /// <summary>
    /// 获取自动排序字段
    /// 对应 PivotField.AutoSortField 属性
    /// </summary>
    string AutoSortField { get; }

    /// <summary>
    /// 获取自动显示类型
    /// 对应 PivotField.AutoShowType 属性
    /// </summary>
    int AutoShowType { get; }

    /// <summary>
    /// 获取所有项是否可见
    /// 对应 PivotField.AllItemsVisible 属性
    /// </summary>
    bool AllItemsVisible { get; }

    /// <summary>
    /// 获取自动显示范围
    /// 对应 PivotField.AutoShowRange 属性
    /// </summary>
    int AutoShowRange { get; }

    /// <summary>
    /// 获取自动显示计数
    /// 对应 PivotField.AutoShowCount 属性
    /// </summary>
    int AutoShowCount { get; }

    /// <summary>
    /// 获取自动显示字段
    /// 对应 PivotField.AutoShowField 属性
    /// </summary>
    string AutoShowField { get; }

    /// <summary>
    /// 获取源标题
    /// 对应 PivotField.SourceCaption 属性
    /// </summary>
    string SourceCaption { get; }

    #endregion

    #region 状态属性
    /// <summary>
    /// 对应 PivotField.EnableItemSelection 属性 
    /// </summary>
    bool EnableItemSelection { get; }

    /// <summary>
    /// 获取数据透视表字段是否为已计算字段
    /// 对应 PivotField.IsCalculated 属性
    /// </summary>
    bool IsCalculated { get; }

    /// <summary>
    /// 对应 PivotField.IsMemberProperty 属性
    /// </summary>
    bool IsMemberProperty { get; }
    #endregion

    #region 图表元素 (子对象)
    /// <summary>
    /// 获取数据透视表字段的数据范围 (如果适用)
    /// 对应 PivotField.DataRange 属性
    /// </summary>
    IExcelRange DataRange { get; }

    /// <summary>
    /// 获取数据透视表字段的标签范围 (如果适用)
    /// 对应 PivotField.LabelRange 属性
    /// </summary>
    IExcelRange LabelRange { get; }
    #endregion


    /// <summary>
    /// 向数据透视表页面字段添加项目
    /// 对应 PivotField.AddPageItem 方法
    /// </summary>
    /// <param name="item">要添加的项目名称</param>
    /// <param name="clearList">是否清除现有项目列表后再添加，null表示使用默认行为</param>
    void AddPageItem(string item, bool? clearList);

    /// <summary>
    /// 删除数据透视表字段
    /// 对应 PivotField.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取数据透视表字段的计算项集合
    /// 对应 PivotField.CalculatedItems 方法
    /// </summary>
    /// <returns>计算项集合对象</returns>
    IExcelCalculatedItems CalculatedItems();

    /// <summary>
    /// 对指定数据字段的值进行自动排序
    /// 对应 PivotField.AutoSort 方法
    /// </summary>
    /// <param name="order">排序顺序，1表示升序，2表示降序</param>
    /// <param name="field">用于排序的数据字段名称</param>
    void AutoSort(int order, string field);

    /// <summary>
    /// 对指定数据字段的值进行扩展自动排序
    /// 对应 PivotField.AutoSortEx 方法
    /// </summary>
    /// <param name="order">排序顺序，1表示升序，2表示降序</param>
    /// <param name="field">用于排序的数据字段名称</param>
    /// <param name="pivotLine">用于排序的数据透视表线对象</param>
    /// <param name="customSubtotal">自定义小计选项</param>
    void AutoSortEx(int order, string field, IExcelPivotLine? pivotLine = null, bool? customSubtotal = null);


    /// <summary>
    /// 根据指定数据字段的值自动显示最前或最后几个项目
    /// 对应 PivotField.AutoShow 方法
    /// </summary>
    /// <param name="type">自动显示类型，1表示显示最大值，2表示显示最小值</param>
    /// <param name="range">自动显示范围，1表示向下，2表示向内</param>
    /// <param name="count">要显示的项目数量</param>
    /// <param name="field">用于确定显示哪些项目的字段名称</param>
    void AutoShow(int type, int range, int count, string field);

    /// <summary>
    /// 钻取到指定的数据透视表字段
    /// 对应 PivotField.DrillTo 方法
    /// </summary>
    /// <param name="field">要钻取到的字段名称</param>
    void DrillTo(string field);

    /// <summary>
    /// 清除手动筛选器
    /// 对应 PivotField.ClearManualFilter 方法
    /// </summary>
    void ClearManualFilter();

    /// <summary>
    /// 清除所有筛选器
    /// 对应 PivotField.ClearAllFilters 方法
    /// </summary>
    void ClearAllFilters();

    /// <summary>
    /// 清除值筛选器
    /// 对应 PivotField.ClearValueFilters 方法
    /// </summary>
    void ClearValueFilters();

    /// <summary>
    /// 清除标签筛选器
    /// 对应 PivotField.ClearLabelFilters 方法
    /// </summary>
    void ClearLabelFilters();

    /// <summary>
    /// 获取指定索引的数据透视表项目
    /// 对应 PivotField.PivotItems 方法
    /// </summary>
    /// <param name="index">项目索引</param>
    /// <returns>数据透视表项目对象</returns>
    [ReturnValueConvert]
    IExcelPivotItems PivotItems(int index);

    /// <summary>
    /// 获取具有指定名称的数据透视表项目
    /// 对应 PivotField.PivotItems 方法
    /// </summary>
    /// <param name="name">项目名称</param>
    /// <returns>数据透视表项目对象</returns>
    [ReturnValueConvert]
    IExcelPivotItems PivotItems(string name);

    /// <summary>
    /// 获取数据透视表项目集合
    /// 对应 PivotField.PivotItems 方法
    /// </summary>
    /// <returns>数据透视表项目集合对象</returns>
    [ReturnValueConvert]
    IExcelPivotItem PivotItems();
}
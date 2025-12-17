//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel PivotTable 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PivotTable 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPivotTable : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取或设置数据透视表的名称
    /// 对应 PivotTable.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取数据透视表的父对象 (通常是 Worksheet)
    /// 对应 PivotTable.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取数据透视表所在的Application对象
    /// 对应 PivotTable.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取数据透视表的源数据缓存
    /// 对应 PivotTable.PivotCache 属性
    /// </summary>
    IExcelPivotCache? PivotCache();

    /// <summary>
    /// 获取或设置数据透视表的源数据
    /// 对应 PivotTable.SourceData 属性 (只读)
    /// </summary>
    object? SourceData { get; }

    /// <summary>
    /// 获取数据透视表的版本
    /// 对应 PivotTable.Version 属性
    /// </summary>
    XlPivotTableVersionList Version { get; }
    #endregion

    #region 数据和字段
    /// <summary>
    /// 获取数据透视表的列范围
    /// </summary>
    IExcelRange? ColumnRange { get; }

    /// <summary>
    /// 获取数据透视表的数据标签范围
    /// </summary>
    IExcelRange? DataLabelRange { get; }

    /// <summary>
    /// 获取或设置内部详细信息
    /// </summary>
    string InnerDetail { get; set; }

    /// <summary>
    /// 获取数据透视表的页面范围
    /// </summary>
    IExcelRange? PageRange { get; }

    /// <summary>
    /// 获取数据透视表的页面范围单元格
    /// </summary>
    IExcelRange? PageRangeCells { get; }

    /// <summary>
    /// 获取数据透视表的行范围
    /// </summary>
    IExcelRange? RowRange { get; }

    /// <summary>
    /// 获取数据透视表的公式集合
    /// </summary>
    IExcelPivotFormulas? PivotFormulas { get; }

    /// <summary>
    /// 获取数据透视表的多维数据集字段
    /// </summary>
    IExcelCubeFields? CubeFields { get; }

    /// <summary>
    /// 获取数据透视表的数据透视字段
    /// </summary>
    IExcelPivotField? DataPivotField { get; }

    /// <summary>
    /// 获取计算成员集合
    /// </summary>
    IExcelCalculatedMembers? CalculatedMembers { get; }

    /// <summary>
    /// 获取数据透视表的列轴
    /// </summary>
    IExcelPivotAxis? PivotColumnAxis { get; }

    /// <summary>
    /// 获取数据透视表的行轴
    /// </summary>
    IExcelPivotAxis? PivotRowAxis { get; }

    /// <summary>
    /// 获取活动筛选器集合
    /// </summary>
    IExcelPivotFilters? ActiveFilters { get; }

    /// <summary>
    /// 获取切片器集合
    /// </summary>
    IExcelSlicers? Slicers { get; }

    /// <summary>
    /// 获取数据透视表更改列表
    /// </summary>
    IExcelPivotTableChangeList? ChangeList { get; }

    /// <summary>
    /// 获取数据透视图表
    /// </summary>
    IExcelShape? PivotChart { get; }

    /// <summary>
    /// 获取或设置替代文本
    /// </summary>
    string AlternativeText { get; set; }

    /// <summary>
    /// 获取或设置摘要信息
    /// </summary>
    string Summary { get; set; }

    /// <summary>
    /// 获取或设置是否为集合显示视觉总计
    /// </summary>
    bool VisualTotalsForSets { get; set; }

    /// <summary>
    /// 获取或设置是否显示值行
    /// </summary>
    bool ShowValuesRow { get; set; }

    /// <summary>
    /// 获取或设置筛选器中是否包含计算成员
    /// </summary>
    bool CalculatedMembersInFilters { get; set; }

    /// <summary>
    /// 获取或设置是否处于网格拖放区域模式
    /// </summary>
    bool InGridDropZones { get; set; }

    /// <summary>
    /// 获取或设置是否显示钻取指示器
    /// </summary>
    bool ShowDrillIndicators { get; set; }

    /// <summary>
    /// 获取或设置打印时是否显示钻取指示器
    /// </summary>
    bool PrintDrillIndicators { get; set; }

    /// <summary>
    /// 获取或设置是否显示成员属性工具提示
    /// </summary>
    bool DisplayMemberPropertyTooltips { get; set; }

    /// <summary>
    /// 获取或设置是否显示上下文工具提示
    /// </summary>
    bool DisplayContextTooltips { get; set; }

    /// <summary>
    /// 获取或设置紧凑布局行缩进
    /// </summary>
    int CompactRowIndent { get; set; }

    /// <summary>
    /// 获取或设置默认布局行类型
    /// </summary>
    XlLayoutRowType LayoutRowDefault { get; set; }

    /// <summary>
    /// 获取或设置是否显示字段标题
    /// </summary>
    bool DisplayFieldCaptions { get; set; }

    /// <summary>
    /// 获取或设置是否查看计算成员
    /// </summary>
    bool ViewCalculatedMembers { get; set; }

    /// <summary>
    /// 获取或设置是否显示空行
    /// </summary>
    bool DisplayEmptyRow { get; set; }

    /// <summary>
    /// 获取或设置是否显示空列
    /// </summary>
    bool DisplayEmptyColumn { get; set; }

    /// <summary>
    /// 获取或设置是否显示来自OLAP的单元格背景
    /// </summary>
    bool ShowCellBackgroundFromOLAP { get; set; }

    /// <summary>
    /// 获取或设置是否显示即时项目
    /// </summary>
    bool DisplayImmediateItems { get; set; }

    /// <summary>
    /// 获取或设置是否显示页面多项标签
    /// </summary>
    bool ShowPageMultipleItemLabel { get; set; }

    /// <summary>
    /// 获取或设置是否启用视觉总计
    /// </summary>
    bool VisualTotals { get; set; }

    /// <summary>
    /// 获取或设置是否启用字段列表
    /// </summary>
    bool EnableFieldList { get; set; }

    /// <summary>
    /// 获取值
    /// </summary>
    string Value { get; }

    /// <summary>
    /// 获取MDX查询字符串
    /// </summary>
    string MDX { get; }

    /// <summary>
    /// 获取刷新日期
    /// </summary>
    DateTime RefreshDate { get; }

    /// <summary>
    /// 获取刷新者姓名
    /// </summary>
    string RefreshName { get; }

    /// <summary>
    /// 获取或设置是否保存数据
    /// </summary>
    bool SaveData { get; set; }

    /// <summary>
    /// 获取或设置缓存索引
    /// </summary>
    int CacheIndex { get; set; }

    /// <summary>
    /// 获取或设置是否显示错误字符串
    /// </summary>
    bool DisplayErrorString { get; set; }

    /// <summary>
    /// 获取或设置是否显示空值字符串
    /// </summary>
    bool DisplayNullString { get; set; }

    /// <summary>
    /// 获取或设置是否启用钻取功能
    /// </summary>
    bool EnableDrilldown { get; set; }

    /// <summary>
    /// 获取或设置是否启用字段对话框
    /// </summary>
    bool EnableFieldDialog { get; set; }

    /// <summary>
    /// 获取或设置错误字符串
    /// </summary>
    string ErrorString { get; set; }

    /// <summary>
    /// 获取或设置是否手动更新
    /// </summary>
    bool ManualUpdate { get; set; }

    /// <summary>
    /// 获取或设置是否合并标签
    /// </summary>
    bool MergeLabels { get; set; }

    /// <summary>
    /// 获取或设置空值字符串
    /// </summary>
    string NullString { get; set; }

    /// <summary>
    /// 获取或设置是否对隐藏的页面项目进行小计
    /// </summary>
    bool SubtotalHiddenPageItems { get; set; }

    /// <summary>
    /// 获取或设置页面字段顺序
    /// </summary>
    int PageFieldOrder { get; set; }

    /// <summary>
    /// 获取或设置页面字段样式
    /// </summary>
    string PageFieldStyle { get; set; }

    /// <summary>
    /// 获取或设置页面字段换行计数
    /// </summary>
    int PageFieldWrapCount { get; set; }

    /// <summary>
    /// 获取或设置是否保留格式
    /// </summary>
    bool PreserveFormatting { get; set; }

    /// <summary>
    /// 获取或设置数据透视表选择
    /// </summary>
    string PivotSelection { get; set; }

    /// <summary>
    /// 获取或设置选择模式
    /// </summary>
    XlPTSelectionMode SelectionMode { get; set; }

    /// <summary>
    /// 获取或设置标签
    /// </summary>
    string Tag { get; set; }

    /// <summary>
    /// 获取或设置腾出空间的样式
    /// </summary>
    string VacatedStyle { get; set; }

    /// <summary>
    /// 获取或设置是否打印标题
    /// </summary>
    bool PrintTitles { get; set; }

    /// <summary>
    /// 获取或设置总计名称
    /// </summary>
    string GrandTotalName { get; set; }

    /// <summary>
    /// 获取或设置是否使用小型网格
    /// </summary>
    bool SmallGrid { get; set; }

    /// <summary>
    /// 获取或设置是否在每页打印时重复项目
    /// </summary>
    bool RepeatItemsOnEachPrintedPage { get; set; }

    /// <summary>
    /// 获取或设置总计注释
    /// </summary>
    bool TotalsAnnotation { get; set; }

    /// <summary>
    /// 获取或设置标准数据透视表选择
    /// </summary>
    string PivotSelectionStandard { get; set; }

    /// <summary>
    /// 获取或设置表格样式2
    /// </summary>
    object TableStyle2 { get; set; }

    /// <summary>
    /// 获取或设置是否显示表格样式行标题
    /// </summary>
    bool ShowTableStyleRowHeaders { get; set; }

    /// <summary>
    /// 获取或设置是否显示表格样式列标题
    /// </summary>
    bool ShowTableStyleColumnHeaders { get; set; }

    /// <summary>
    /// 获取或设置是否允许多重筛选
    /// </summary>
    bool AllowMultipleFilters { get; set; }

    /// <summary>
    /// 获取或设置紧凑布局行标题
    /// </summary>
    string CompactLayoutRowHeader { get; set; }

    /// <summary>
    /// 获取或设置紧凑布局列标题
    /// </summary>
    string CompactLayoutColumnHeader { get; set; }

    /// <summary>
    /// 获取或设置字段列表是否按升序排序
    /// </summary>
    bool FieldListSortAscending { get; set; }

    /// <summary>
    /// 获取或设置是否使用自定义列表排序
    /// </summary>
    bool SortUsingCustomLists { get; set; }

    /// <summary>
    /// 获取或设置位置
    /// </summary>
    string Location { get; set; }

    /// <summary>
    /// 获取或设置是否启用回写功能
    /// </summary>
    bool EnableWriteback { get; set; }

    /// <summary>
    /// 获取或设置分配方式
    /// </summary>
    XlAllocation Allocation { get; set; }

    /// <summary>
    /// 获取或设置分配值
    /// </summary>
    XlAllocationValue AllocationValue { get; set; }

    /// <summary>
    /// 获取或设置分配方法
    /// </summary>
    XlAllocationMethod AllocationMethod { get; set; }

    /// <summary>
    /// 获取或设置分配权重表达式
    /// </summary>
    string AllocationWeightExpression { get; set; }
    #endregion

    #region 格式和布局
    /// <summary>
    /// 获取或设置数据透视表的表格样式
    /// 对应 PivotTable.TableStyle2 属性
    /// </summary>
    string TableStyle { get; set; }

    /// <summary>
    /// 获取或设置数据透视表是否显示行条纹
    /// 对应 PivotTable.ShowTableStyleRowStripes 属性
    /// </summary>
    bool ShowTableStyleRowStripes { get; set; }

    /// <summary>
    /// 获取或设置数据透视表是否显示列条纹
    /// 对应 PivotTable.ShowTableStyleColumnStripes 属性
    /// </summary>
    bool ShowTableStyleColumnStripes { get; set; }


    /// <summary>
    /// 获取或设置数据透视表是否显示末列特殊样式
    /// 对应 PivotTable.ShowTableStyleLastColumn 属性
    /// </summary>
    bool ShowTableStyleLastColumn { get; set; }

    /// <summary>
    /// 获取或设置数据透视表是否显示行总计
    /// 对应 PivotTable.RowGrand 属性
    /// </summary>
    bool RowGrand { get; set; }

    /// <summary>
    /// 获取或设置数据透视表是否显示列总计
    /// 对应 PivotTable.ColumnGrand 属性
    /// </summary>
    bool ColumnGrand { get; set; }

    /// <summary>
    /// 获取或设置数据透视表在刷新或移动字段时是否自动设置格式
    /// 对应 PivotTable.HasAutoFormat 属性
    /// </summary>
    bool HasAutoFormat { get; set; }
    #endregion

    #region 状态属性 


    bool EnableWizard { get; set; }

    bool EnableDataValueEditing { get; set; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 获取数据透视表中的特定字段
    /// </summary>
    /// <param name="Index">要获取的字段索引或名称</param>
    /// <returns>对应的数据透视表字段，如果未找到则返回null</returns>
    [ReturnValueConvert]
    IExcelPivotField? PivotFields(int Index);

    /// <summary>
    /// 获取数据透视表中的特定字段
    /// </summary>
    /// <param name="Index">要获取的字段索引或名称</param>
    /// <returns>对应的数据透视表字段，如果未找到则返回null</returns>
    [ReturnValueConvert]
    IExcelPivotField? PivotFields(string Index);

    /// <summary>
    /// 获取数据透视表中所有字段的集合
    /// </summary>
    /// <returns>包含所有数据透视表字段的集合，如果出现错误则返回null</returns>
    [ReturnValueConvert]
    IExcelPivotFields? PivotFields();

    /// <summary>
    /// 获取数据透视表中所有计算字段的集合
    /// 对应 PivotTable.CalculatedFields 属性
    /// </summary>
    /// <returns>包含所有计算字段的集合，如果出现错误则返回null</returns>
    IExcelCalculatedFields? CalculatedFields();

    #endregion

    #region 数据透视表操作  

    /// <summary>
    /// 更新数据透视表 (通常与 Refresh 同义)
    /// </summary>
    void Update();

    /// <summary>
    /// 删除当前应用于 PivotTable的所有筛选器。
    /// </summary>
    void ClearAllFilters();

    /// <summary>
    /// 清除 PivotTable包括删除所有字段以及删除应用于PivotTables的所有筛选和排序。
    /// </summary>
    void ClearTable();

    /// <summary>
    /// 更改指定 PivotTable的连接。
    /// </summary>
    /// <param name="conn">必需的 <see cref="IExcelWorkbookConnection"/> 对象，该对象表示新的 PivotTable 连接。</param>
    void ChangeConnection(IExcelWorkbookConnection conn);

    /// <summary>
    /// 基于 OLAP 数据源对数据透视表的数据源执行提交操作。
    /// </summary>
    void CommitChanges();

    /// <summary>
    /// 用于将PivotTable转换为多维数据集公式。
    /// </summary>
    /// <param name="convertFilters"></param>
    void ConvertToFormulas(bool convertFilters);

    /// <summary>
    /// 创建数据透视表的多维数据集文件，该数据透视表与联机分析处理 (OLAP) 数据源相连接。
    /// </summary>
    /// <param name="file">要创建的多维数据集文件的名称。 如果该文件已存在，则将被覆盖。</param>
    /// <param name="measures">度量的唯一名称的数组</param>
    /// <param name="levels">字符串数组。 每个数组项都是一个唯一的级别名称。</param>
    /// <param name="members"> 字符串数组的数组。 元素对应于数组中 Levels 表示的层次结构。</param>
    /// <param name="properties"></param>
    /// <returns></returns>
    string CreateCubeFile(string file, string[]? measures = null, string[]? levels = null, string[]? members = null, bool? properties = null);

    /// <summary>
    /// 放弃基于 OLAP 数据源的数据透视表中经过编辑的单元格中的所有更改。
    /// </summary>
    void DiscardChanges();

    /// <summary>
    /// 向下钻取到基于 OLAP 或 PowerPivot 的多维数据集层次结构中的数据。
    /// </summary>
    /// <param name="pivotItem">执行向下钻取的成员。</param>
    /// <param name="pivotLine">指定操作起始成员所在的数据透视表中的行。</param>
    void DrillDown(IExcelPivotItem pivotItem, IExcelRange? pivotLine);

    /// <summary>
    /// 钻取到指定的多维数据集字段
    /// </summary>
    /// <param name="pivotItem">数据透视表项</param>
    /// <param name="cubeField">多维数据集字段</param>
    /// <param name="pivotLine">数据透视表行范围</param>
    void DrillTo(IExcelPivotItem pivotItem, IExcelCubeField cubeField, IExcelRange? pivotLine);

    /// <summary>
    /// 向上钻取到指定层级
    /// </summary>
    /// <param name="PivotItem">数据透视表项</param>
    /// <param name="pivotLine">数据透视表行范围</param>
    /// <param name="levelUniqueName">层级唯一名称</param>
    void DrillUp(IExcelPivotItem PivotItem, IExcelRange? pivotLine, object? levelUniqueName);

    /// <summary>
    /// 设置数据透视表的格式类型
    /// </summary>
    /// <param name="Format">透视表格式类型</param>
    void Format(XlPivotFormatType Format);

    /// <summary>
    /// 更改数据透视表的数据缓存
    /// </summary>
    /// <param name="pivotCache">数据透视表缓存名称</param>
    void ChangePivotCache(string pivotCache);

    /// <summary>
    /// 更改数据透视表的数据缓存
    /// </summary>
    /// <param name="pivotCache">数据透视表缓存对象</param>
    void ChangePivotCache(IExcelPivotCache pivotCache);

    /// <summary>
    /// 返回一个 Range 对象，该对象带有数据透视表中数据项的相关信息。
    /// </summary>
    IExcelRange? GetPivotData(string? dataField = null,
        string? field1 = null, string? Item1 = null, string? field2 = null,
        string? item2 = null, string? field3 = null, string? item3 = null,
        string? field4 = null, string? item4 = null, string? field5 = null,
        string? item5 = null, string? field6 = null, string? item6 = null,
        string? field7 = null, string? item7 = null, string? field8 = null,
        string? item8 = null, string? field9 = null, string? item9 = null,
        string? field10 = null, string? item10 = null, string? field11 = null,
        string? item11 = null, string? field12 = null, string? item12 = null,
        string? field13 = null, string? item13 = null, string? field14 = null,
        string? item14 = null);

    /// <summary>
    /// 获取数据透视表中指定行列位置的数值单元格
    /// </summary>
    /// <param name="rowline">行索引，从0开始计数</param>
    /// <param name="columnline">列索引，从0开始计数</param>
    /// <returns>指定位置的数据透视表数值单元格对象</returns>
    IExcelPivotValueCell? PivotValueCell(int? rowline, int? columnline);

    /// <summary>
    /// 分离工作表上创建数据透视表的计算项和计算字段的列表。
    /// </summary>
    void ListFormulas();

    /// <summary>
    /// 用源数据刷新数据透视表。
    /// </summary>
    /// <returns></returns>
    bool? RefreshTable();

    /// <summary>
    /// 为处于回写模式的数据透视表中所有编辑过的单元格，从数据源检索当前值。
    /// </summary>
    void RefreshDataSourceValues();

    /// <summary>
    /// 设置是否为指定数据透视表中的所有数据透视字段重复项目标签。
    /// </summary>
    /// <param name="repeat">指定是否要对指定数据透视表中的所有透视字段重复项目标签。</param>
    void RepeatAllLabels(XlPivotFieldRepeatLabels repeat);

    /// <summary>
    /// 用于同时设置所有现有 PivotField的布局选项。
    /// </summary>
    /// <param name="rowLayout">可以是 xlCompactRow、 xlTabularRow 或 xlOutlineRow。</param>
    void RowAxisLayout(XlLayoutRowType rowLayout);

    /// <summary>
    /// 为页字段中的每个数据项创建新的数据透视表。
    /// </summary>
    /// <param name="pageField"></param>
    /// <returns></returns>
    object? ShowPages(string pageField);

    /// <summary>
    /// 选定数据透视表的一部分。
    /// </summary>
    /// <param name="name">所选内容</param>
    /// <param name="mode">结构化选择模式。</param>
    /// <param name="useStandardName">如果为 True，则表示将在其他位置运行录制的宏。</param>
    void PivotSelect(string name, XlPTSelectionMode mode = XlPTSelectionMode.xlDataAndLabel, bool? useStandardName = null);

    /// <summary>
    /// 更改所有现有 PivotFields的分类汇总位置。
    /// </summary>
    /// <param name="location">可以是 xlAtTop 或 xlAtBottom。</param>
    void SubtotalLocation(XlSubtototalLocationType location);

    /// <summary>
    /// 从指定的数据透视表单元格返回数据。
    /// </summary>
    /// <param name="name">描述数据透视表中的单个单元格</param>
    /// <returns></returns>
    double? GetData(string name);

    /// <summary>
    /// 在基于 OLAP 数据源的数据透视表中的所有已编辑的单元格上执行回写操作。
    /// </summary>
    void AllocateChanges();

    /// <summary>
    /// 向数据透视表添加数据字段
    /// </summary>
    /// <param name="field">要添加的字段，可以是字段名称或字段索引</param>
    /// <param name="caption">字段的显示标题，如果为null则使用默认标题</param>
    /// <param name="function">聚合函数名称，如"Sum"、"Count"等</param>
    /// <returns>添加的数据透视表字段对象，如果添加失败则返回null</returns>
    IExcelPivotField? AddDataField(object field, string? caption, string? function);

    /// <summary>
    /// 向数据透视表或数据透视图中添加行字段、列字段和页字段。
    /// </summary>
    /// <param name="rowFields">可选 对象。 指定要添加为行或要添加到类别轴的字段名称(或字段名称数组)。</param>
    /// <param name="columnFields">可选 对象。 指定字段名称 (或字段名称数组)， 添加为列或要添加到序列轴。</param>
    /// <param name="pageFields">可选 对象。 指定字段名称(或字段名称数组) ，添加为页或要添加到页面区域。</param>
    /// <param name="addToTable">可选 对象。 仅适用于数据透视表。 如果为 True，则将指定的字段添加到报表中（不替换现有字段）。 如果为 False，则用新的字段替换现有的字段。 默认值为 False。</param>
    /// <returns></returns>
    object? AddFields(object? rowFields, object? columnFields, object? pageFields, bool? addToTable);

    #endregion
}

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
    /// 获取数据透视表的数据主体区域 (不包括页字段报告筛选器)
    /// 对应 PivotTable.DataBodyRange 属性
    /// </summary>
    IExcelRange? DataBodyRange { get; }

    /// <summary>
    /// 获取数据透视表的整个表格区域 (包括页字段报告筛选器)
    /// 对应 PivotTable.TableRange1 属性
    /// </summary>
    IExcelRange? TableRange1 { get; }

    /// <summary>
    /// 获取数据透视表的第二区域 (如果有页字段报告筛选器，则包含这些字段)
    /// 对应 PivotTable.TableRange2 属性
    /// </summary>
    IExcelRange? TableRange2 { get; }

    /// <summary>
    /// 获取数据透视表的页字段 (报告筛选器) 集合
    /// 对应 PivotTable.PageFields 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelPivotFields? PageFields { get; }

    /// <summary>
    /// 获取数据透视表的行字段集合
    /// 对应 PivotTable.RowFields 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelPivotFields? RowFields { get; }

    /// <summary>
    /// 获取数据透视表的列字段集合
    /// 对应 PivotTable.ColumnFields 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelPivotFields? ColumnFields { get; }

    /// <summary>
    /// 获取数据透视表的数据字段集合
    /// 对应 PivotTable.DataFields 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelPivotFields? DataFields { get; }

    /// <summary>
    /// 获取数据透视表的可见数据字段集合
    /// 对应 PivotTable.VisibleFields 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelPivotFields? VisibleFields { get; }

    /// <summary>
    /// 获取数据透视表的隐藏数据字段集合
    /// 对应 PivotTable.HiddenFields 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelPivotFields? HiddenFields { get; }
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

    void DrillTo(IExcelPivotItem pivotItem, IExcelCubeField cubeField, IExcelRange? pivotLine);

    void DrillUp(IExcelPivotItem PivotItem, IExcelRange? pivotLine, object? levelUniqueName);

    void Format(XlPivotFormatType Format);

    void ChangePivotCache(string pivotCache);

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

    IExcelPivotValueCell PivotValueCell(int? rowline, int? columnline);

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
    object ShowPages(string pageField);

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

    IExcelPivotField? AddDataField(object field, string? caption, string? function);

    /// <summary>
    /// 向数据透视表或数据透视图中添加行字段、列字段和页字段。
    /// </summary>
    /// <param name="rowFields">可选 对象。 指定要添加为行或要添加到类别轴的字段名称(或字段名称数组)。</param>
    /// <param name="columnFields">可选 对象。 指定字段名称 (或字段名称数组)， 添加为列或要添加到序列轴。</param>
    /// <param name="pageFields">可选 对象。 指定字段名称(或字段名称数组) ，添加为页或要添加到页面区域。</param>
    /// <param name="addToTable">可选 对象。 仅适用于数据透视表。 如果为 True，则将指定的字段添加到报表中（不替换现有字段）。 如果为 False，则用新的字段替换现有的字段。 默认值为 False。</param>
    /// <returns></returns>
    object AddFields(object? rowFields, object? columnFields, object? pageFields, bool? addToTable);

    #endregion
}

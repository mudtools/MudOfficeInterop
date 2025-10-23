//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Data;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 工作表中一个单元格区域的包装器接口，提供对单元格区域的各种操作和属性访问功能。
/// </summary>
public interface IExcelRange : ICoreRange<IExcelRange>, IDisposable
{
    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="column">列号</param>
    /// <returns>单元格对象</returns>
    IExcelRange? this[int? row, int? column] { get; }

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="rowAddress">行地址</param>
    /// <param name="columnAddress">列地址</param>
    /// <returns>单元格对象</returns>
    IExcelRange? this[string? rowAddress, string? columnAddress] { get; }

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="address">地址</param>
    IExcelRange? this[string address] { get; }

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <returns>单元格对象</returns>
    IExcelRange? this[int row] { get; }



    /// <summary>
    /// 从DataTable复制数据到Excel工作表
    /// </summary>
    /// <param name="dataTable">数据表</param>
    /// <param name="startCell">起始单元格</param>
    /// <param name="fieldNames">是否包含字段名</param>
    /// <returns>是否操作成功</returns>
    bool CopyFromDataTable(DataTable dataTable, string startCell = "A1", bool fieldNames = true);


    /// <summary>
    /// 在区域中替换指定内容
    /// </summary>
    /// <param name="what">要替换的内容</param>
    /// <param name="replacement">替换内容</param>
    /// <param name="lookAt">匹配方式</param>
    /// <param name="searchOrder">查找顺序</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="matchByte">是否匹配字节</param>
    /// <param name="searchFormat">查找格式</param>
    /// <param name="replaceFormat">替换格式</param>
    /// <returns>是否成功替换</returns>
    bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte, object searchFormat, object replaceFormat);

    /// <summary>
    /// 为区域创建分类汇总
    /// </summary>
    /// <param name="groupBy">分组依据列</param>
    /// <param name="function">汇总函数</param>
    /// <param name="totalList">需要汇总的列列表</param>
    /// <param name="replace">是否替换现有汇总</param>
    /// <param name="pageBreaks">是否在组间插入分页符</param>
    /// <param name="summaryBelowData">汇总行位置</param>
    void Subtotal(int groupBy, XlConsolidationFunction function, object totalList, bool replace, bool pageBreaks, XlSummaryRow summaryBelowData);

}


/// <summary>
/// 获取地址的参数选项
/// </summary>
public class AddressOptions
{
    public bool? RowAbsolute { get; set; }
    public bool? ColumnAbsolute { get; set; }
    public XlReferenceStyle ReferenceStyle { get; set; } = XlReferenceStyle.xlA1;
    public bool? External { get; set; }
    public bool? RelativeTo { get; set; }

    public AddressOptions()
    {
    }

    public AddressOptions(bool absolute)
    {
        RowAbsolute = absolute;
        ColumnAbsolute = absolute;
    }
}
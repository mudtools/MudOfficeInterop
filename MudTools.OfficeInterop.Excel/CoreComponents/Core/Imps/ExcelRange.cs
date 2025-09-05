//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Data;

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// 初始化Excel范围对象
/// </summary>
/// <param name="range">Excel原生Range对象</param>
/// <exception cref="ArgumentNullException">range参数为空时抛出</exception>
internal class ExcelRange(MsExcel.Range? range) : CoreRange<ExcelRange, IExcelRange>(range),
    IExcelRange, IExcelRows, IExcelColumns, IDisposable
{

    #region 构造函数

    public ExcelRange() : this(null)
    {

    }
    #endregion

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="column">列号</param>
    /// <returns>单元格对象</returns>
    public IExcelRange? this[int? row, int? column]
    {
        get
        {
            return _range?.Item[row.ComArgsVal(i => i > 0), column.ComArgsVal(i => i > 0)] is MsExcel.Range rang ? new ExcelRange(rang) : null;
        }
    }

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="rowAddress">行地址</param>
    /// <param name="columnAddress">列地址</param>
    /// <returns>单元格对象</returns>
    public IExcelRange? this[string? rowAddress, string? columnAddress]
    {
        get
        {
            return _range?.Item[rowAddress.ComArgsVal(), columnAddress.ComArgsVal()] is MsExcel.Range rang ? new ExcelRange(rang) : null;
        }
    }

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="address">地址</param>
    public IExcelRange? this[string address]
    {
        get
        {
            if (string.IsNullOrWhiteSpace(address))
                throw new ArgumentException("地址不能为空");
            return _range?[address] is MsExcel.Range rang ? new ExcelRange(rang) : null;
        }
    }


    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <returns>单元格对象</returns>
    public IExcelRange? this[int row] => this[row, null];



    /// <summary>
    /// 从DataTable复制数据到Excel工作表
    /// </summary>
    /// <param name="dataTable">数据表</param>
    /// <param name="startCell">起始单元格</param>
    /// <param name="fieldNames">是否包含字段名</param>
    /// <returns>是否操作成功</returns>
    public bool CopyFromDataTable(
        DataTable dataTable,
        string startCell = "A1",
        bool fieldNames = true)
    {
        try
        {
            MsExcel.Range startRange = _range.Worksheet.Range[startCell];
            int startRow = startRange.Row;
            int startCol = startRange.Column;

            // 写入列标题
            if (fieldNames)
            {
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    var cell = _range.Worksheet.Cells[startRow, startCol + i] as MsExcel.Range; ;
                    cell.Value = dataTable.Columns[i].ColumnName;
                }
                startRow++; // 数据从下一行开始
            }

            // 写入数据
            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    var cell = _range.Worksheet.Cells[startRow + row, startCol + col] as MsExcel.Range;
                    cell.Value = dataTable.Rows[row][col];
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            throw new ExcelOperationException($"从DataTable复制数据失败: {ex.Message}", ex);
        }
    }


    /// <summary>
    /// 替换内容
    /// </summary>
    /// <param name="what">要替换的内容</param>
    /// <param name="replacement">替换后的内容</param>
    /// <param name="lookAt">匹配方式</param>
    /// <param name="searchOrder">搜索顺序</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="matchByte">双字节匹配</param>
    /// <param name="searchFormat">搜索格式</param>
    /// <param name="replaceFormat">替换格式</param>
    /// <returns>是否发生替换</returns>
    public bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte, object searchFormat, object replaceFormat)
    {
        return InternalRange.Replace(
            What: what,
            Replacement: replacement,
            LookAt: lookAt,
            SearchOrder: searchOrder,
            MatchCase: matchCase,
            MatchByte: matchByte,
            SearchFormat: searchFormat,
            ReplaceFormat: replaceFormat);
    }


    /// <summary>
    /// 创建分类汇总
    /// </summary>
    /// <param name="groupBy">分组依据的列索引</param>
    /// <param name="function">汇总函数</param>
    /// <param name="totalList">汇总列索引数组</param>
    /// <param name="replace">是否替换现有分类汇总</param>
    /// <param name="pageBreaks">是否分页</param>
    /// <param name="summaryBelowData">汇总位置</param>
    public void Subtotal(int groupBy, int function, object totalList, bool replace, bool pageBreaks, int summaryBelowData)
    {
        InternalRange.Subtotal(
            GroupBy: groupBy,
            Function: (MsExcel.XlConsolidationFunction)function,
            TotalList: totalList,
            Replace: replace,
            PageBreaks: pageBreaks,
            SummaryBelowData: (MsExcel.XlSummaryRow)summaryBelowData);
    }
}

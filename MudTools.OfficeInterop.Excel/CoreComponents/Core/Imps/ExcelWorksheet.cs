//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

partial class ExcelWorksheet
{

    /// <summary>
    /// 获取工作表中指定范围的区域对象
    /// </summary>
    /// <param name="cell1">起始单元格</param>
    /// <param name="cell2">结束单元格（可选）</param>
    /// <returns>区域对象</returns>
    public IExcelRange? Range(object? cell1, object? cell2 = null)
    {
        if (_worksheet == null) return null;

        try
        {
            if (cell1 is ExcelRange range1)
                cell1 = range1.InternalRange;
            if (cell2 is ExcelRange range2)
                cell2 = range2.InternalRange;

            cell1 ??= System.Type.Missing;
            cell2 ??= System.Type.Missing;

            var range = _worksheet.Range[cell1, cell2]; ;
            return range != null ? new ExcelRange(range) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="column">列号</param>
    /// <returns>单元格对象</returns>
    public IExcelRange? this[int row, int column]
    {
        get
        {
            if (_worksheet == null) return null;

            var r = _worksheet.Cells[row, column] is MsExcel.Range range ? new ExcelRange(range) : null;
            if (r != null) _disposableList.Add(r);
            return r;
        }
    }

    public IExcelRange? this[string address]
    {
        get
        {
            if (_worksheet == null) return null;
            var range = _worksheet.Range[address];
            var r = range != null ? new ExcelRange(range) : null;
            if (r != null) _disposableList.Add(r);
            return r;
        }
    }

    public IExcelRange? this[string begin, string end] => this[$"{begin}:{end}"];
}
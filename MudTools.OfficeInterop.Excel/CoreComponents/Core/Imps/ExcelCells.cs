//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelCells : CoreRange<ExcelRange, IExcelRange>, IExcelCells
{
    #region 构造函数
    /// <summary>
    /// 初始化Excel范围对象
    /// </summary>
    /// <param name="range">Excel原生Range对象</param>
    /// <exception cref="ArgumentNullException">range参数为空时抛出</exception>
    internal ExcelCells(MsExcel.Range range) : base(range)
    {
    }

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
            return _range[address] is MsExcel.Range rang ? new ExcelRange(rang) : null;
        }
    }


    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <returns>单元格对象</returns>
    public IExcelRange? this[int row] => this[row, null];


    ~ExcelCells()
    {
        Dispose(false);
    }

    #endregion
}

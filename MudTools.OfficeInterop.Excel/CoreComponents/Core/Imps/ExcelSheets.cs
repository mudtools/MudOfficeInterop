//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelSheets : ExcelCommonSheets, IExcelSheets
{
    private MsExcel.Sheets _worksheets;
    /// <summary>
    /// 用于记录此类型运行时日志的 logger 实例。
    /// </summary>
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelSheets));
    internal ExcelSheets(MsExcel.Sheets sheets)
    {
        _worksheets = sheets ?? throw new ArgumentNullException(nameof(sheets));
    }

    #region 基础属性
    public override int Count => _worksheets.Count;

    /// <summary>
    /// 获取指定索引的工作表对象
    /// </summary>
    /// <param name="index">工作表索引（从1开始）</param>
    /// <returns>工作表对象</returns>
    public override IExcelCommonSheet? this[int index]
    {
        get
        {
            if (_worksheets == null || index < 1 || index > Count)
                return null;

            try
            {
                var sheet = _worksheets[index];
                if (sheet != null && sheet is MsExcel.Worksheet worksheet)
                    return new ExcelWorksheet(worksheet);
                if (sheet != null && sheet is MsExcel.Chart chart)
                    return new ExcelChart(chart);
                return null;
            }
            catch (Exception ex)
            {
                log.Warn($"Failed to retrieve sheet at index {index}: {ex.Message}");
                return null;
            }
        }
    }

    /// <summary>
    /// 获取指定名称的工作表对象
    /// </summary>
    /// <param name="name">工作表名称</param>
    /// <returns>工作表对象</returns>
    public override IExcelCommonSheet? this[string name]
    {
        get
        {
            if (_worksheets == null || string.IsNullOrEmpty(name))
                return null;

            try
            {
                var sheet = _worksheets[name];
                if (sheet != null && sheet is MsExcel.Worksheet worksheet)
                    return new ExcelWorksheet(worksheet);
                if (sheet != null && sheet is MsExcel.Chart chart)
                    return new ExcelChart(chart);
                return null;
            }
            catch (Exception ex)
            {
                log.Warn($"Failed to retrieve sheet with name '{name}': {ex.Message}");
                return null;
            }
        }
    }


    public override object Parent => _worksheets.Parent;

    protected override object NativeSheets => _worksheets;

    public override IExcelApplication Application => new ExcelApplication(_worksheets.Application);
    #endregion

    #region 创建和添加
    public override IExcelWorksheet? AddSheet(
        IExcelCommonSheet? before = null,
        IExcelCommonSheet? after = null,
        int? count = 1)
    {
        object? beforeObj = before switch
        {
            ExcelWorksheet ws => ws.Worksheet,
            ExcelChart chart => chart._chart,
            _ => Type.Missing
        };

        object? afterObj = after switch
        {
            ExcelWorksheet ws => ws.Worksheet,
            ExcelChart chart => chart._chart,
            _ => Type.Missing
        };

        object result = _worksheets.Add(
                        beforeObj,
                        afterObj,
                        count.ComArgsVal(),
                        MsExcel.XlSheetType.xlWorksheet);
        if (result is MsExcel.Worksheet workSheet)
            return new ExcelWorksheet(workSheet);
        return null;
    }

    public IExcelChart? AddChart(
       IExcelCommonSheet? before = null,
       IExcelCommonSheet? after = null,
       int? count = 1)
    {
        object? beforeObj = before switch
        {
            ExcelWorksheet ws => ws.Worksheet,
            ExcelChart chart => chart._chart,
            _ => Type.Missing
        };

        object? afterObj = after switch
        {
            ExcelWorksheet ws => ws.Worksheet,
            ExcelChart chart => chart._chart,
            _ => Type.Missing
        };

        object result = _worksheets.Add(
                        beforeObj,
                        afterObj,
                        count.ComArgsVal(),
                        MsExcel.XlSheetType.xlChart);
        if (result is MsExcel.Chart workSheet)
            return new ExcelChart(workSheet);
        return null;
    }

    public override IExcelCommonSheet? Add(
                                    IExcelCommonSheet? before = null,
                                    IExcelCommonSheet? after = null,
                                    int? count = 1,
                                    XlSheetType? type = null)
    {
        object? beforeObj = before switch
        {
            ExcelWorksheet ws => ws.Worksheet,
            ExcelChart chart => chart._chart,
            _ => Type.Missing
        };

        object? afterObj = after switch
        {
            ExcelWorksheet ws => ws.Worksheet,
            ExcelChart chart => chart._chart,
            _ => Type.Missing
        };

        object result = _worksheets.Add(
                        beforeObj,
                        afterObj,
                        count.ComArgsVal(),
                        type.ComArgsConvert(x => (MsExcel.XlSheetType)(int)x).ComArgsVal());

        return result switch
        {
            MsExcel.Worksheet ws => new ExcelWorksheet(ws),
            MsExcel.Chart nchart => new ExcelChart(nchart),
            _ => null
        };
    }

    public override IExcelCommonSheet? CreateFromTemplate(
                                        string templatePath,
                                        string sheetName,
                                        IExcelCommonSheet? before = null,
                                        IExcelCommonSheet? after = null)
    {
        if (_worksheets == null || string.IsNullOrEmpty(templatePath))
            return null;

        try
        {
            object beforeObj = before is ExcelWorksheet ws ? ws.Worksheet : Type.Missing;
            object afterObj = after is ExcelWorksheet aw ? aw.Worksheet : Type.Missing;

            if (_worksheets.Add(beforeObj, afterObj, Type.Missing, templatePath) is not MsExcel.Worksheet worksheet)
                return null;

            var excelWorksheet = new ExcelWorksheet(worksheet);
            if (!string.IsNullOrEmpty(sheetName))
                excelWorksheet.Name = sheetName;

            return excelWorksheet;
        }
        catch (Exception ex)
        {
            log.Warn($"Failed to create sheet from template '{templatePath}': {ex.Message}");
            return null;
        }
    }
    #endregion

    #region 操作方法
    /// <summary>
    /// 将此 Sheets 集合中的所有工作表复制到指定位置。
    /// 这是对 Microsoft.Office.Interop.Excel.Sheets.Copy 方法的封装。
    /// </summary>
    /// <param name="beforeSheet">
    /// 指定应在哪个工作表之前放置复制的工作表。
    /// 如果为 null，则不指定此参数。
    /// </param>
    /// <param name="afterSheet">
    /// 指定应在哪个工作表之后放置复制的工作表。
    /// 如果为 null，则不指定此参数。
    /// </param>
    /// <exception cref="System.InvalidOperationException">
    /// 如果内部的 Sheets 对象为 null。
    /// </exception>
    /// <exception cref="System.Runtime.InteropServices.COMException">
    /// 如果与 Excel 的交互失败（例如，参数无效，工作表被保护），可能会抛出 COM 异常。
    /// </exception>
    /// <remarks>
    /// 如果 beforeSheet 和 afterSheet 都为 null，则 Excel 通常会创建一个新工作簿来容纳复制的工作表。
    /// 如果同时指定了 beforeSheet 和 afterSheet，行为可能不确定（通常 After 会被忽略）。
    /// </remarks>
    public void CopyTo(
        IExcelCommonSheet? beforeSheet = null,
        IExcelCommonSheet? afterSheet = null)
    {
        // 检查内部对象是否为 null
        if (_worksheets == null)
        {
            log.Error("Underlying Sheets object is null in CopyTo method.");
            throw new InvalidOperationException("Cannot copy Sheets: underlying Interop Sheets object is null.");
        }

        try
        {
            object? beforeObj = beforeSheet switch
            {
                ExcelWorksheet ws => ws.Worksheet,
                ExcelChart chart => chart._chart,
                _ => Type.Missing
            };

            object? afterObj = afterSheet switch
            {
                ExcelWorksheet ws => ws.Worksheet,
                ExcelChart chart => chart._chart,
                _ => Type.Missing
            };
            _worksheets.Copy(beforeObj, afterObj);
        }
        catch (COMException comEx)
        {
            // 记录或重新抛出特定的 COM 异常
            log.Error($"COM Exception in CopyTo method: {comEx.Message}", comEx);
            // 可以选择包装异常或直接重新抛出
            throw; // 重新抛出，让调用者处理
        }
        catch (Exception ex) // 捕获其他可能的异常
        {
            log.Error($"General Exception in CopyTo method: {ex.Message}", ex);
            throw new InvalidOperationException("Failed to copy sheets.", ex);
        }
    }

    /// <summary>
    /// 将此 Sheets 集合中的所有工作表移动到指定位置。
    /// 这是对 Microsoft.Office.Interop.Excel.Sheets.Move 方法的封装。
    /// </summary>
    /// <param name="beforeSheet">
    /// 指定应在哪个工作表之前放置移动的工作表。
    /// 如果为 null，则不指定此参数。
    /// </param>
    /// <param name="afterSheet">
    /// 指定应在哪个工作表之后放置移动的工作表。
    /// 如果为 null，则不指定此参数。
    /// </param>
    /// <exception cref="System.InvalidOperationException">
    /// 如果内部的 Sheets 对象为 null。
    /// </exception>
    /// <exception cref="System.Runtime.InteropServices.COMException">
    /// 如果与 Excel 的交互失败（例如，参数无效，工作表被保护），可能会抛出 COM 异常。
    /// </exception>
    /// <remarks>
    /// 如果 beforeSheet 和 afterSheet 都为 null，行为可能不确定（可能移动到新工作簿或失败）。
    /// 如果同时指定了 beforeSheet 和 afterSheet，行为可能不确定（通常 After 会被忽略）。
    /// </remarks>
    public void MoveTo(IExcelCommonSheet? beforeSheet = null, IExcelCommonSheet? afterSheet = null)
    {
        // 检查内部对象是否为 null
        if (_worksheets == null)
        {
            log.Error("Underlying Sheets object is null in CopyTo method.");
            throw new InvalidOperationException("Cannot move Sheets: underlying Interop Sheets object is null.");
        }

        try
        {
            object? beforeObj = beforeSheet switch
            {
                ExcelWorksheet ws => ws.Worksheet,
                ExcelChart chart => chart._chart,
                _ => Type.Missing
            };

            object? afterObj = afterSheet switch
            {
                ExcelWorksheet ws => ws.Worksheet,
                ExcelChart chart => chart._chart,
                _ => Type.Missing
            };

            // 调用 Interop 的 Move 方法
            _worksheets.Move(beforeObj, afterObj);
        }
        catch (COMException comEx)
        {
            log.Error($"COM Exception in MoveTo method: {comEx.Message}", comEx);
            throw;
        }
        catch (Exception ex)
        {
            log.Error($"General Exception in MoveTo method: {ex.Message}", ex);
            throw new InvalidOperationException("Failed to move sheets.", ex);
        }
    }


    /// <summary>
    /// 将指定区域的内容和格式填充到此 Sheets 集合中所有工作表的对应区域。
    /// 这是对 Microsoft.Office.Interop.Excel.Sheets.FillAcrossSheets 方法的封装。
    /// </summary>
    /// <param name="sourceRange">
    /// 代表要填充的源区域的 ExcelRange 对象。
    /// </param>
    /// <param name="fillType">
    /// 指定要填充的内容类型（全部、仅内容、仅格式）。
    /// </param>
    /// <exception cref="System.ArgumentNullException">
    /// 如果 sourceRange 为 null。
    /// </exception>
    /// <exception cref="System.InvalidOperationException">
    /// 如果内部的 Sheets 对象为 null。
    /// </exception>
    /// <exception cref="System.Runtime.InteropServices.COMException">
    /// 如果与 Excel 的交互失败（例如，源区域无效，工作表被保护），可能会抛出 COM 异常。
    /// </exception>
    public void FillAcrossSheets(IExcelRange sourceRange, XlFillWith? fillType = XlFillWith.xlFillWithAll)
    {
        // 检查内部对象是否为 null
        if (_worksheets == null)
        {
            log.Error("Underlying Sheets object is null in CopyTo method.");
            throw new InvalidOperationException("Cannot fill across sheets: underlying Interop Sheets object is null.");
        }

        // 检查源区域参数
        if (sourceRange == null)
        {
            throw new ArgumentNullException(nameof(sourceRange), "Source range cannot be null.");
        }

        try
        {
            MsExcel.Range interopRange = ((ExcelRange)sourceRange).InternalRange;
            _worksheets.FillAcrossSheets(interopRange,
                fillType.EnumConvert(MsExcel.XlFillWith.xlFillWithAll));
        }
        catch (COMException comEx)
        {
            // 记录或重新抛出特定的 COM 异常
            log.Error($"COM Exception in FillAcrossSheets method: {comEx.Message}", comEx);
            throw; // 重新抛出，让调用者处理
        }
        catch (Exception ex) // 捕获其他可能的异常
        {
            log.Error($"General Exception in FillAcrossSheets method: {ex.Message}", ex);
            throw new InvalidOperationException("Failed to fill across sheets.", ex);
        }
    }

    public void Clear()
    {
        if (_worksheets == null || Count <= 1) return;

        try
        {
            // 从后往前删，避免索引偏移
            for (int i = Count; i >= 2; i--)
            {
                Delete(i);
            }
        }
        catch (Exception ex)
        {
            log.Warn($"Failed to clear sheets: {ex.Message}");
        }
    }

    public override void Delete(int index)
    {
        if (_worksheets == null || index < 1 || index > Count) return;

        try
        {
            if (_worksheets[index] is MsExcel.Worksheet ws)
                ws.Delete();
            else if (_worksheets[index] is MsExcel.Chart chart)
                chart.Delete();
        }
        catch (Exception ex)
        {
            log.Warn($"Failed to delete sheet at index {index}: {ex.Message}");
        }
    }

    public override void Delete(string name)
    {
        if (_worksheets == null || string.IsNullOrEmpty(name)) return;

        try
        {
            if (_worksheets[name] is MsExcel.Worksheet ws)
                ws.Delete();
            else if (_worksheets[name] is MsExcel.Chart chart)
                chart.Delete();
        }
        catch (Exception ex)
        {
            log.Warn($"Failed to delete sheet named '{name}': {ex.Message}");
        }
    }

    public override void Delete(IExcelCommonSheet sheet)
    {
        if (sheet == null || _worksheets == null) return;

        try
        {
            if (sheet is ExcelWorksheet ws)
                ws.Worksheet.Delete();
            else if (sheet is ExcelChart chart)
                chart._chart.Delete();
        }
        catch (Exception ex)
        {
            log.Warn($"Failed to delete provided sheet: {ex.Message}");
        }
    }

    /// <summary>
    /// 选择多个工作表
    /// </summary>
    /// <param name="worksheetNames">工作表名称数组</param>
    public override void Select(params string[] worksheetNames)
    {
        if (_worksheets == null || worksheetNames == null || worksheetNames.Length == 0)
            return;

        try
        {
            _worksheets.Select(worksheetNames.Cast<object>().ToArray());
        }
        catch (Exception ex)
        {
            log.Warn($"Failed to select worksheets: {ex.Message}");
        }
    }
    #endregion  

    #region 导出和导入
    public int ExportToFolder(string folderPath, string fileFormat = "xlsx", string prefix = "sheet_")
    {
        if (!Directory.Exists(folderPath))
            return 0;

        int count = 0;
        foreach (var sheet in this)
        {
            string fileName = Path.Combine(folderPath, $"{prefix}{sheet.Name}.{fileFormat}");
            try
            {
                sheet.SaveAs(fileName);
                count++;
            }
            catch (Exception ex)
            {
                log.Warn($"Failed to export sheet '{sheet.Name}' to {fileName}: {ex.Message}");
            }
        }
        return count;
    }
    #endregion

    #region IEnumerable<IExcelWorksheet> Support
    public override IEnumerator<IExcelCommonSheet> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            var sheet = this[i];
            if (sheet != null)
                yield return sheet;
        }
    }
    #endregion

    /// <summary>
    /// 计算所有工作表
    /// </summary>
    public override void Calculate()
    {
        if (_worksheets == null) return;

        foreach (var sheet in this)
        {
            if (sheet is IExcelWorksheet ws)
            {
                try
                {
                    ws.Calculate();
                }
                catch (Exception ex)
                {
                    log.Warn($"Failed to calculate sheet '{ws.Name}': {ex.Message}");
                }
            }
        }
    }

    /// <summary>
    /// 打印所有工作表
    /// </summary>
    /// <param name="preview">是否打印预览</param>
    public override void PrintOutAll(bool preview = false)
    {
        if (_worksheets == null) return;

        try
        {
            if (preview)
                _worksheets.PrintPreview();
            else
                _worksheets.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                     Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
        catch (Exception ex)
        {
            log.Warn($"Failed to print all sheets: {ex.Message}");
        }
    }

    /// <summary>
    /// 刷新所有工作表
    /// </summary>
    public override void RefreshAll()
    {
        if (_worksheets == null) return;

        foreach (var sheet in this)
        {
            if (sheet is IExcelWorksheet ws)
            {
                try
                {
                    ws.Recalculate();
                }
                catch (Exception ex)
                {
                    log.Warn($"Failed to refresh sheet '{ws.Name}': {ex.Message}");
                }
            }
        }
    }

    /// <summary>
    /// 获取活动工作表
    /// </summary>
    /// <returns>活动工作表对象</returns>
    public override IExcelCommonSheet? ActiveWorksheet
    {
        get
        {
            try
            {
                if (_worksheets?.Parent is not MsExcel.Workbook workbook) return null;

                var active = workbook.ActiveSheet;
                return active switch
                {
                    MsExcel.Worksheet ws => new ExcelWorksheet(ws),
                    MsExcel.Chart chart => new ExcelChart(chart),
                    _ => null
                };
            }
            catch (Exception ex)
            {
                log.Warn($"Failed to get active sheet: {ex.Message}");
                return null;
            }
        }
    }

    /// <summary>
    /// 隐藏所有工作表
    /// </summary>
    public void HideAll()
    {
        if (_worksheets == null || Count == 0) return;

        try
        {
            for (int i = 2; i <= Count; i++) // 保留第一个可见
            {
                if (this[i] is IExcelCommonSheet sheet)
                    sheet.Visible = XlSheetVisibility.xlSheetHidden;
            }
        }
        catch (Exception ex)
        {
            log.Warn($"Failed to hide all sheets: {ex.Message}");
        }
    }

    /// <summary>
    /// 显示所有工作表
    /// </summary>
    public void ShowAll()
    {
        if (_worksheets == null || Count == 0) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                if (this[i] is IExcelCommonSheet sheet)
                    sheet.Visible = XlSheetVisibility.xlSheetVisible;
            }
        }
        catch (Exception ex)
        {
            log.Warn($"Failed to show all sheets: {ex.Message}");
        }
    }

    #region IDisposable Support
    protected override void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                if (_worksheets != null)
                {
                    Marshal.ReleaseComObject(_worksheets);
                    _worksheets = null;
                }
            }
            _disposedValue = true;
        }
        base.Dispose(disposing);
    }

    ~ExcelSheets()
    {
        Dispose(false);
    }
    #endregion
}

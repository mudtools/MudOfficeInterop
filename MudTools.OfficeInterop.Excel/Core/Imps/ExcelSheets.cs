//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
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
    public override IExcelWorksheet? this[int index]
    {
        get
        {
            if (_worksheets == null || index < 1 || index > Count)
                return null;

            try
            {
                return _worksheets[index] is MsExcel.Worksheet worksheet ? new ExcelWorksheet(worksheet) : null;
            }
            catch
            {
                return null;
            }
        }
    }
    public override object Parent => _worksheets.Parent;

    protected override object NativeSheets => _worksheets;

    public override IExcelApplication Application => new ExcelApplication(_worksheets.Application);
    #endregion

    #region 创建和添加
    public IExcelWorksheet? Add(
        object? before = null,
        object? after = null,
        object? count = null,
        object? type = null)
    {
        object comBefore = before ?? System.Type.Missing;
        object comAfter = after ?? System.Type.Missing;
        object comCount = count ?? System.Type.Missing;
        object comType = type ?? System.Type.Missing;

        object newSheet = _worksheets.Add(comBefore, comAfter, comCount, comType);

        if (newSheet is MsExcel.Worksheet newWs)
        {
            return new ExcelWorksheet(newWs);
        }
        return null;
    }

    public IExcelWorksheet? CreateFromTemplate(string templatePath, string sheetName, object? before = null, object? after = null)
    {
        if (_worksheets == null || string.IsNullOrEmpty(templatePath))
            return null;

        try
        {
            ExcelWorksheet? beforeSheet = before as ExcelWorksheet;
            ExcelWorksheet? afterSheet = after as ExcelWorksheet;


            if (_worksheets.Add(
                beforeSheet?.Worksheet,
                afterSheet?.Worksheet,
                Type.Missing,
                templatePath
            ) is MsExcel.Worksheet worksheet)
            {
                ExcelWorksheet excelWorksheet = new ExcelWorksheet(worksheet);
                if (!string.IsNullOrEmpty(sheetName))
                {
                    excelWorksheet.Name = sheetName;
                }
                return excelWorksheet;
            }
            return null;
        }
        catch
        {
            return null;
        }
    }
    #endregion

    #region 查找和筛选
    public IExcelWorksheet[] GetVisibleSheets()
    {
        List<IExcelWorksheet> results = new List<IExcelWorksheet>();
        for (int i = 1; i <= Count; i++)
        {
            var worksheet = this[i];
            if (worksheet != null && worksheet.Visible == XlSheetVisibility.xlSheetVisible)
            {
                results.Add(worksheet);
            }
        }
        return results.ToArray();
    }

    public IExcelWorksheet[] GetHiddenSheets()
    {
        List<IExcelWorksheet> results = new List<IExcelWorksheet>();
        for (int i = 1; i <= Count; i++)
        {
            var worksheet = this[i];
            if (worksheet != null && worksheet.Visible == XlSheetVisibility.xlSheetHidden)
            {
                results.Add(worksheet);
            }
        }
        return results.ToArray();
    }

    public IExcelWorksheet[] GetVeryHiddenSheets()
    {
        List<IExcelWorksheet> results = new List<IExcelWorksheet>();
        for (int i = 1; i <= Count; i++)
        {
            IExcelWorksheet worksheet = this[i];
            if (worksheet != null && worksheet.Visible == XlSheetVisibility.xlSheetVeryHidden)
            {
                results.Add(worksheet);
            }
        }
        return results.ToArray();
    }

    public IExcelWorksheet[] GetProtectedSheets()
    {
        List<IExcelWorksheet> results = new List<IExcelWorksheet>();
        for (int i = 1; i <= Count; i++)
        {
            IExcelWorksheet? worksheet = this[i];
            if (worksheet != null && worksheet.IsProtected)
            {
                results.Add(worksheet);
            }
        }
        return results.ToArray();
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
    public void CopyTo(object? beforeSheet = null, object? afterSheet = null)
    {
        // 检查内部对象是否为 null
        if (_worksheets == null)
        {
            log.Error("Underlying Sheets object is null in CopyTo method.");
            throw new InvalidOperationException("Cannot copy Sheets: underlying Interop Sheets object is null.");
        }

        try
        {
            object? interopBefore = (beforeSheet as ExcelWorksheet)?.Worksheet ?? (beforeSheet as ExcelChart)?._chart ?? beforeSheet;
            object? interopAfter = (afterSheet as ExcelWorksheet)?.Worksheet ?? (afterSheet as ExcelChart)?._chart ?? afterSheet;

            _worksheets.Copy(interopBefore, interopAfter);
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
    public void MoveTo(object? beforeSheet = null, object? afterSheet = null)
    {
        // 检查内部对象是否为 null
        if (_worksheets == null)
        {
            log.Error("Underlying Sheets object is null in CopyTo method.");
            throw new InvalidOperationException("Cannot move Sheets: underlying Interop Sheets object is null.");
        }

        try
        {
            // 处理可选参数 (同 CopyTo)
            object? interopBefore = (beforeSheet as ExcelWorksheet)?.Worksheet ?? (beforeSheet as ExcelChart)?._chart ?? beforeSheet;
            object? interopAfter = (afterSheet as ExcelWorksheet)?.Worksheet ?? (afterSheet as ExcelChart)?._chart ?? afterSheet;


            // 调用 Interop 的 Move 方法
            _worksheets.Move(interopBefore, interopAfter);
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
    public void FillAcrossSheets(IExcelRange sourceRange, XlFillWith fillType)
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
            _worksheets.FillAcrossSheets(interopRange, (MsExcel.XlFillWith)fillType);
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
        for (int i = Count; i >= 1; i--)
        {
            try
            {
                // 注意：不能删除所有工作表
                if (Count > 1)
                {
                    Delete(i);
                }
            }
            catch
            {
                // 处理错误
            }
        }
    }

    public override void Delete(int index)
    {
        try
        {
            MsExcel.Worksheet? sheetToDelete = _worksheets[index] as MsExcel.Worksheet;
            sheetToDelete?.Delete();
        }
        catch
        {
        }
    }

    public override void Delete(string name)
    {
        try
        {
            MsExcel.Worksheet? sheetToDelete = _worksheets[name] as MsExcel.Worksheet;
            sheetToDelete?.Delete();
        }
        catch
        {
            // 忽略异常
        }
    }

    public override void Delete(IExcelWorksheet sheet)
    {
        if (sheet is ExcelWorksheet realSheet)
        {
            try
            {
                realSheet.Worksheet.Delete();
            }
            catch { }
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
            object[] sheets = new object[worksheetNames.Length];
            for (int i = 0; i < worksheetNames.Length; i++)
            {
                sheets[i] = worksheetNames[i];
            }
            _worksheets.Select(sheets);
        }
        catch
        {

        }
    }
    #endregion  

    #region 导出和导入
    public int ExportToFolder(string folderPath, string fileFormat = "xlsx", string prefix = "sheet_")
    {
        if (!Directory.Exists(folderPath)) return 0;

        int count = 0;
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                IExcelWorksheet? sheet = this[i];
                string fileName = Path.Combine(folderPath, $"{prefix}{sheet?.Name}.{fileFormat}");
                sheet?.SaveAs(fileName);
                count++;
            }
            catch
            {
                // Log error
            }
        }
        return count;
    }

    public IExcelWorksheet[]? ImportFromFile(string filename, string sheetName = "", object? before = null, object? after = null)
    {
        if (_worksheets == null || string.IsNullOrEmpty(filename))
            return null;

        MsExcel.Application? excelApp = null;
        MsExcel.Workbook? sourceWorkbook = null;

        try
        {
            // 获取当前Excel应用程序实例
            string progId = "Excel.Application";
#if ns21
            excelApp = (MsExcel.Application)ComInterop.GetActiveObject(progId);
#else
            excelApp = (MsExcel.Application)Marshal.GetActiveObject(progId);
#endif
            bool newApp = false;

            // 如果没有运行的Excel实例则创建新实例
            if (excelApp == null)
            {
                excelApp = new MsExcel.Application();
                newApp = true;
                excelApp.Visible = false;  // 隐藏Excel窗口
            }

            // 打开源工作簿
            sourceWorkbook = excelApp.Workbooks.Open(
                Filename: filename,
                UpdateLinks: false,
                ReadOnly: true,
                IgnoreReadOnlyRecommended: true
            );

            // 获取第一个工作表
            MsExcel.Worksheet? sourceSheet =
                (MsExcel.Worksheet)sourceWorkbook.Sheets[1];

            // 确定目标位置
            MsExcel.Worksheet? targetSheet = null;
            if (before != null)
                targetSheet = ((ExcelWorksheet)before).Worksheet;
            else if (after != null)
                targetSheet = ((ExcelWorksheet)after).Worksheet;

            // 复制工作表到当前工作簿
            sourceSheet.Copy(
                Before: targetSheet != null && before != null ? targetSheet : Type.Missing,
                After: targetSheet != null && after != null ? targetSheet : Type.Missing
            );

            // 获取新添加的工作表（最后一个工作表）
            MsExcel.Worksheet newSheet =
                (MsExcel.Worksheet)excelApp.ActiveWorkbook.Sheets[excelApp.ActiveWorkbook.Sheets.Count];

            // 创建包装对象并返回
            return [new ExcelWorksheet(newSheet)];
        }
        catch (COMException)
        {
            // Excel未安装或COM异常
            return null;
        }
        catch (System.IO.FileNotFoundException)
        {
            // 文件不存在
            return null;
        }
        catch (System.UnauthorizedAccessException)
        {
            // 文件访问权限问题
            return null;
        }
        finally
        {
            // 清理资源
            if (sourceWorkbook != null)
            {
                sourceWorkbook.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceWorkbook);
            }

            if (excelApp != null)
            {
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }
    }
    #endregion

    #region IEnumerable<IExcelWorksheet> Support
    public override IEnumerator<IExcelWorksheet> GetEnumerator()
    {
        for (int i = 1; i <= _worksheets.Count; i++)
        {
            yield return new ExcelWorksheet(_worksheets[i] as MsExcel.Worksheet);
        }
    }
    #endregion

    /// <summary>
    /// 计算所有工作表
    /// </summary>
    public override void Calculate()
    {
        if (_worksheets == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    this[i]?.Calculate();
                }
                catch
                {
                    // 忽略单个工作表计算异常
                }
            }
        }
        catch
        {
            // 忽略计算过程中的异常
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
            {
                _worksheets.PrintPreview();
            }
            else
            {
                _worksheets.PrintOut(
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing
                );
            }
        }
        catch
        {
            // 忽略打印过程中的异常
        }
    }

    /// <summary>
    /// 刷新所有工作表
    /// </summary>
    public override void RefreshAll()
    {
        if (_worksheets == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    this[i]?.Recalculate();
                }
                catch
                {
                    // 忽略单个工作表刷新异常
                }
            }
        }
        catch
        {
            // 忽略刷新过程中的异常
        }
    }

    /// <summary>
    /// 获取活动工作表
    /// </summary>
    /// <returns>活动工作表对象</returns>
    public override IExcelWorksheet ActiveWorksheet
    {
        get
        {
            try
            {
                MsExcel.Workbook? wb = _worksheets?.Parent as MsExcel.Workbook;
                if (wb == null)
                    return null;
                var activeSheet = wb.ActiveSheet as MsExcel.Worksheet;
                return activeSheet != null ? new ExcelWorksheet(activeSheet) : null;
            }
            catch
            {
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
            // 保留第一个工作表可见，隐藏其余工作表
            for (int i = 2; i <= Count; i++)
            {
                try
                {
                    this[i].Visible = XlSheetVisibility.xlSheetHidden;
                }
                catch
                {
                    // 忽略单个工作表隐藏异常
                }
            }
        }
        catch
        {
            // 忽略隐藏过程中的异常
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
                try
                {
                    this[i].Visible = XlSheetVisibility.xlSheetVisible;
                }
                catch
                {
                    // 忽略单个工作表显示异常
                }
            }
        }
        catch
        {
            // 忽略显示过程中的异常
        }
    }

    #region IDisposable Support
    protected override void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // 释放托管状态(托管对象)
            }

            if (_worksheets != null)
            {
                try
                {
                    while (Marshal.ReleaseComObject(_worksheets) > 0) { }
                }
                catch
                {
                    // 忽略释放过程中可能发生的异常
                }
                _worksheets = null;
            }

            _disposedValue = true;
        }
        base.Dispose(disposing);
    }

    ~ExcelSheets()
    {

    }
    #endregion
}

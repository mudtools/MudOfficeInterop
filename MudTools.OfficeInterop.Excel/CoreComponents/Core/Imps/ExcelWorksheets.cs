//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelWorksheets : ExcelCommonSheets, IExcelWorksheets
{
    /// <summary>
    /// 底层的 COM Worksheets 集合对象
    /// </summary>
    private MsExcel.Sheets? _worksheets;
    private DisposableList _disposables = [];
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelWorksheets));

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelWorksheets 实例
    /// </summary>
    /// <param name="worksheets">底层的 COM Worksheets 集合对象</param>
    internal ExcelWorksheets(MsExcel.Sheets worksheets)
    {
        _worksheets = worksheets;
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected override void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _disposables?.Dispose();
            if (_worksheets != null)
                Marshal.ReleaseComObject(_worksheets);
            _worksheets = null;
        }
        _disposedValue = true;
        base.Dispose(disposing);
    }

    ~ExcelWorksheets()
    {
        Dispose(false);
    }
    #endregion

    #region 基础属性

    /// <summary>
    /// 获取工作表集合中的工作表数量
    /// </summary>
    public override int Count => _worksheets?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的工作表对象
    /// </summary>
    /// <param name="index">工作表索引（从1开始）</param>
    /// <returns>工作表对象</returns>
    public IExcelWorksheet? this[int index]
    {
        get
        {
            if (_worksheets == null || index < 1 || index > Count)
                return null;

            try
            {
                var sheet = _worksheets[index];
                ExcelWorksheet? excelWorksheet = null;
                if (sheet != null && sheet is MsExcel.Worksheet worksheet)
                {
                    excelWorksheet = new ExcelWorksheet(worksheet);
                    _disposables.Add(excelWorksheet);
                }
                return excelWorksheet;
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
    public IExcelWorksheet? this[string name]
    {
        get
        {
            if (_worksheets == null || string.IsNullOrEmpty(name))
                return null;

            try
            {
                var sheet = _worksheets.Item[name];
                ExcelWorksheet? excelWorksheet = null;
                if (sheet != null && sheet is MsExcel.Worksheet worksheet)
                {
                    excelWorksheet = new ExcelWorksheet(worksheet);
                    _disposables.Add(excelWorksheet);
                }
                return null;
            }
            catch (Exception ex)
            {
                log.Warn($"Failed to retrieve sheet with name '{name}': {ex.Message}");
                return null;
            }
        }
    }
    public override IEnumerable<IExcelComSheet> Items()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    protected override IExcelComSheet? ItemByIndex(int index)
    {
        return this[index];
    }


    /// <summary>
    /// 获取工作表集合所在的父对象
    /// </summary>
    public override object? Parent => _worksheets?.Parent;

    protected override object? NativeSheets => _worksheets;

    /// <summary>
    /// 获取工作表集合所在的Application对象
    /// </summary>
    public override IExcelApplication? Application
    {
        get
        {
            MsExcel.Application? application = _worksheets?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }
    #endregion

    #region 创建和添加
    public override IExcelWorksheet? AddSheet(
       IExcelComSheet? before = null,
       IExcelComSheet? after = null,
       int? count = 1)
    {
        object? beforeObj = before switch
        {
            ExcelWorksheet ws => ws.InternalComObject,
            ExcelChart chart => chart._chart,
            _ => Type.Missing
        };

        object? afterObj = after switch
        {
            ExcelWorksheet ws => ws.InternalComObject,
            ExcelChart chart => chart._chart,
            _ => Type.Missing
        };

        object? result = _worksheets?.Add(
                        beforeObj,
                        afterObj,
                        count.ComArgsVal(),
                        MsExcel.XlSheetType.xlWorksheet);
        if (result is MsExcel.Worksheet workSheet)
            return new ExcelWorksheet(workSheet);
        return null;
    }

    /// <summary>
    /// 向工作簿添加新的工作表
    /// </summary>
    /// <param name="before">添加到指定工作表之前</param>
    /// <param name="after">添加到指定工作表之后</param>
    /// <param name="count">添加的工作表数量</param>
    /// <param name="type">工作表类型</param>
    /// <returns>新创建的工作表对象</returns>
    public override IExcelComSheet? Add(
                                IExcelComSheet? before = null,
                                IExcelComSheet? after = null,
                                int? count = 1,
                                XlSheetType? type = null)
    {
        if (_worksheets == null)
            return null;

        try
        {
            object? beforeObj = before switch
            {
                ExcelWorksheet ws => ws.InternalComObject,
                ExcelChart chart => chart._chart,
                _ => Type.Missing
            };

            object? afterObj = after switch
            {
                ExcelWorksheet ws => ws.InternalComObject,
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
        catch (Exception ex)
        {
            log?.Warn($"Failed to add worksheet: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// 批量添加工作表
    /// </summary>
    /// <param name="names">工作表名称数组</param>
    /// <param name="before">添加到指定工作表之前</param>
    /// <param name="after">添加到指定工作表之后</param>
    /// <returns>成功添加的工作表数量</returns>
    public int AddRange(string[] names, IExcelWorksheet? before = null, IExcelWorksheet? after = null)
    {
        if (_worksheets == null || names == null || names.Length == 0)
            return 0;

        int successCount = 0;
        foreach (string name in names)
        {
            var worksheet = Add(before, after);
            if (worksheet != null && worksheet.Name != name)
            {
                worksheet.Name = name;
                successCount++;
            }
        }
        return successCount;
    }

    /// <summary>
    /// 基于模板创建工作表
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <param name="name">工作表名称</param>
    /// <param name="before">添加到指定工作表之前</param>
    /// <param name="after">添加到指定工作表之后</param>
    /// <returns>新创建的工作表对象</returns>
    public override IExcelComSheet? CreateFromTemplate(string templatePath, string name = "",
                                            IExcelComSheet? before = null, IExcelComSheet? after = null)
    {
        if (_worksheets == null || string.IsNullOrEmpty(templatePath))
            return null;

        try
        {
            object? beforeObj = before switch
            {
                ExcelWorksheet ws => ws.InternalComObject,
                ExcelChart chart => chart._chart,
                _ => Type.Missing
            };

            object? afterObj = after switch
            {
                ExcelWorksheet ws => ws.InternalComObject,
                ExcelChart chart => chart._chart,
                _ => Type.Missing
            };

            if (_worksheets.Add(beforeObj, afterObj, Type.Missing, templatePath
            ) is MsExcel.Worksheet worksheet)
            {
                var excelWorksheet = new ExcelWorksheet(worksheet);
                if (!string.IsNullOrEmpty(name))
                {
                    excelWorksheet.Name = name;
                }
                return excelWorksheet;
            }
            return null;
        }
        catch (Exception ex)
        {
            log.Warn($"Failed to create worksheet from template '{templatePath}': {ex.Message}");
            return null;
        }
    }

    #endregion

    #region 查找和筛选

    private IEnumerable<IExcelWorksheet> EnumerateWorksheets()
    {
        for (int i = 1; i <= Count; i++)
        {
            if (this[i] is IExcelWorksheet ws)
                yield return ws;
        }
    }

    public IExcelWorksheet[] GetVisibleSheets(XlSheetVisibility visible)
         => EnumerateWorksheets().Where(w => w.Visible == visible).ToArray();

    public IExcelWorksheet[] GetVisibleWorksheets()
        => GetVisibleSheets(XlSheetVisibility.xlSheetVisible);

    public IExcelWorksheet[] GetHiddenWorksheets()
        => GetVisibleSheets(XlSheetVisibility.xlSheetHidden);

    public IExcelWorksheet[] GetProtectedWorksheets()
        => EnumerateWorksheets().Where(w => w.IsProtected).ToArray();

    public IExcelWorksheet[] GetUnprotectedWorksheets()
        => EnumerateWorksheets().Where(w => !w.IsProtected).ToArray();

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除所有工作表（除了第一个）
    /// </summary>
    public void Clear()
    {
        if (_worksheets == null || Count <= 1) return;

        try
        {
            // 从后往前删除，避免索引变化问题
            for (int i = Count; i >= 2; i--)
            {
                var worksheet = _worksheets[i] as MsExcel.Worksheet;
                worksheet.Delete();
            }
        }
        catch (Exception ex)
        {
            log.Warn("Failed to clear worksheets", ex);
        }
    }

    public override void Delete(int index) => DeleteInternal(index);

    public override void Delete(string name) => DeleteInternal(name);

    public override void Delete(IExcelComSheet worksheet)
    {
        if (worksheet is IExcelWorksheet ws)
            ws.Delete();
    }

    private void DeleteInternal(object key)
    {
        if (_worksheets == null) return;

        try
        {
            if (key is int index && index >= 1 && index <= Count)
            {
                if (this[index] is IExcelWorksheet ws) ws.Delete();
            }
            else if (key is string name && !string.IsNullOrEmpty(name))
            {
                if (this[name] is IExcelWorksheet ws) ws.Delete();
            }
        }
        catch (Exception ex)
        {
            log.Warn($"Failed to delete worksheet by {key}: {ex.Message}");
        }
    }

    /// <summary>
    /// 移动工作表
    /// </summary>
    /// <param name="worksheet">要移动的工作表</param>
    /// <param name="before">移动到指定工作表之前</param>
    /// <param name="after">移动到指定工作表之后</param>
    public void Move(IExcelWorksheet worksheet, IExcelWorksheet? before = null, IExcelWorksheet? after = null)
    {
        if (_worksheets == null || worksheet == null)
            return;

        try
        {
            worksheet.Move(before, after);
        }
        catch (Exception ex)
        {
            log.Warn("Failed to move worksheet", ex);
        }
    }

    /// <summary>
    /// 复制工作表
    /// </summary>
    /// <param name="worksheet">要复制的工作表</param>
    /// <param name="before">复制到指定工作表之前</param>
    /// <param name="after">复制到指定工作表之后</param>
    /// <param name="newName">新工作表名称</param>
    /// <returns>复制的工作表对象</returns>
    public IExcelWorksheet? Copy(IExcelWorksheet worksheet, IExcelWorksheet? before = null,
                              IExcelWorksheet? after = null, string newName = "")
    {
        if (_worksheets == null || worksheet == null)
            return null;

        try
        {
            worksheet.Copy(before, after);

            // 获取复制后的工作表
            var copiedWorksheet = this[Count] as ExcelWorksheet;
            if (copiedWorksheet != null && !string.IsNullOrEmpty(newName))
            {
                copiedWorksheet.Name = newName;
            }

            return copiedWorksheet;
        }
        catch (Exception ex)
        {
            log.Warn("Failed to copy worksheet", ex);
        }
        return null;
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
        catch (Exception ex)
        {
            log.Warn("Failed to select worksheets", ex);
        }
    }

    #endregion

    /// <summary>
    /// 按指定顺序排列工作表
    /// </summary>
    /// <param name="names">工作表名称顺序数组</param>
    public void ArrangeInOrder(string[] names)
    {
        if (_worksheets == null || names == null || names.Length == 0)
            return;

        try
        {
            // 按照指定顺序重新排列工作表
            for (int i = 0; i < names.Length; i++)
            {
                if (_worksheets[names[i]] is MsExcel.Worksheet worksheet && worksheet.Index != i + 1)
                {
                    if (i == 0)
                        worksheet.Move(_worksheets[1], Type.Missing);
                    else
                        worksheet.Move(Type.Missing, _worksheets[i]);
                }
            }
        }
        catch (Exception ex)
        {
            log.Warn("Failed to arrange worksheets in order", ex);
        }
    }

    /// <summary>
    /// 按名称排序工作表
    /// </summary>
    /// <param name="ascending">是否升序排列</param>
    public void SortByName(bool ascending = true)
    {
        if (_worksheets == null || Count <= 1) return;

        var names = EnumerateWorksheets().Select(w => w.Name).OrderBy(x => x, StringComparer.Ordinal).ToArray();
        if (!ascending)
            Array.Reverse(names);

        ArrangeInOrder(names);
    }

    #region 高级功能

    /// <summary>
    /// 获取活动工作表
    /// </summary>
    /// <returns>活动工作表对象</returns>
    public override IExcelComSheet? ActiveWorksheet
    {
        get
        {
            try
            {
                if (_worksheets?.Parent is not MsExcel.Workbook wb)
                    return null;
                return wb.ActiveSheet is MsExcel.Worksheet activeSheet ? new ExcelWorksheet(activeSheet) : null;
            }
            catch
            {
                return null;
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
    /// 刷新所有工作表
    /// </summary>
    public override void RefreshAll()
    {
        if (_worksheets == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                if (this[i] is IExcelWorksheet worksheet)
                    worksheet.Recalculate();
            }
        }
        catch (Exception ex)
        {
            log.Warn($"Failed to get active sheet: {ex.Message}");
            return;
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
                this[i].Visible = XlSheetVisibility.xlSheetHidden;
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
                this[i].Visible = XlSheetVisibility.xlSheetVisible;
            }
        }
        catch (Exception ex)
        {
            log.Warn($"Failed to show all sheets: {ex.Message}");
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public IEnumerator<IExcelWorksheet> GetEnumerator()
    {
        for (int i = 0; i < Count; i++)
        {
            yield return this[i];
        }
    }
    #endregion
}
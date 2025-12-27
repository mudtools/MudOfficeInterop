//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel Workbooks 集合对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Workbooks 对象的安全访问和资源管理
/// </summary>
internal class ExcelWorkbooks : IExcelWorkbooks
{
    /// <summary>
    /// 底层的 COM Workbooks 集合对象
    /// </summary>
    private MsExcel.Workbooks? _workbooks;
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelWorkbooks));
    private DisposableList _disposables = [];

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelWorkbooks 实例
    /// </summary>
    /// <param name="workbooks">底层的 COM Workbooks 集合对象</param>
    internal ExcelWorkbooks(MsExcel.Workbooks workbooks)
    {
        _workbooks = workbooks;
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放底层COM对象
            _disposables?.Dispose();
            if (_workbooks != null)
                Marshal.ReleaseComObject(_workbooks);
            _workbooks = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    /// <summary>
    /// 获取工作簿集合中的工作簿数量
    /// </summary>
    public int Count => _workbooks?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的工作簿对象
    /// </summary>
    /// <param name="index">工作簿索引（从1开始）</param>
    /// <returns>工作簿对象</returns>
    public IExcelWorkbook? this[int index]
    {
        get
        {
            if (_workbooks == null || index < 1 || index > Count)
                return null;

            try
            {
                var workbook = _workbooks[index];
                var wb = workbook != null ? new ExcelWorkbook(workbook) : null;
                if (wb != null)
                    _disposables.Add(wb);
                return wb;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取指定名称的工作簿对象
    /// </summary>
    /// <param name="name">工作簿名称</param>
    /// <returns>工作簿对象</returns>
    public IExcelWorkbook? this[string name]
    {
        get
        {
            if (_workbooks == null || string.IsNullOrEmpty(name))
                return null;

            try
            {
                IExcelWorkbook[] result = FindByName(name);
                if (result != null && result.Length > 0)
                    return result[0];
                return null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取工作簿集合所在的父对象
    /// </summary>
    public object? Parent => _workbooks?.Parent;

    /// <summary>
    /// 获取工作簿集合所在的Application对象
    /// </summary>
    public IExcelApplication? Application
    {
        get
        {
            var application = _workbooks?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    #endregion

    #region 创建和打开

    /// <inheritdoc/>
    /// <exception cref="InvalidOperationException"></exception>
    public IExcelWorkbook? Open(string filename, int? updateLinks = 0, bool? readOnly = false,
                       int? format = 1, string? password = null, string? writeResPassword = null,
                       bool? ignoreReadOnlyRecommended = false, XlPlatform? origin = null,
                       string? delimiter = ",", bool? editable = null, bool? notify = null,
                       int? converter = null, bool? addToMru = false, bool? local = null,
                        XlCorruptLoad? corruptLoad = XlCorruptLoad.xlNormalLoad)
    {
        if (_workbooks == null || string.IsNullOrEmpty(filename))
            return null;
        try
        {
            var workbook = _workbooks.Open(
                filename, updateLinks.ComArgsVal(), readOnly.ComArgsVal(), format.ComArgsVal(),
                password.ComArgsVal(), writeResPassword.ComArgsVal(), ignoreReadOnlyRecommended.ComArgsVal(),
                origin.ComArgsConvert(d => d.EnumConvert(MsExcel.XlPlatform.xlWindows)),
                delimiter.ComArgsVal(), editable.ComArgsVal(), notify.ComArgsVal(),
                converter.ComArgsVal(), addToMru.ComArgsVal(), local.ComArgsVal(),
                corruptLoad.ComArgsConvert(d => d.EnumConvert(MsExcel.XlCorruptLoad.xlNormalLoad)));

            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
        catch (COMException ce)
        {
            if (ce.ErrorCode == -2147221040)
                throw new InvalidOperationException("Failed to open workbook. The file is corrupted.", ce);
            else
                throw new InvalidOperationException("Failed to open workbook.", ce);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to open workbook.", ex);
        }
    }

    public IExcelWorkbook? Add(XlWBATemplate template)
    {
        if (_workbooks == null)
            return null;
        try
        {
            MsExcel.Workbook workbook = _workbooks.Add(template.EnumConvert(MsExcel.XlWBATemplate.xlWBATWorksheet));
            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
        catch (COMException ce)
        {
            if (ce.ErrorCode == -2147221040)
                throw new InvalidOperationException("Failed to create workbook from template. The file is corrupted.", ce);
            else
                throw new InvalidOperationException("Failed to create workbook from template.", ce);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create workbook from template.", ex);
        }
    }

    /// <summary>
    /// 新建工作簿
    /// </summary>
    /// <param name="template">模板文件路径</param>
    /// <returns>新建的工作簿对象</returns>
    public IExcelWorkbook? Add(string template = "")
    {
        if (_workbooks == null)
            return null;

        try
        {
            MsExcel.Workbook workbook;
            if (string.IsNullOrEmpty(template))
            {
                workbook = _workbooks.Add();
            }
            else
            {
                workbook = _workbooks.Add(template);
            }

            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
        catch (COMException ce)
        {
            if (ce.ErrorCode == -2147221040)
                throw new InvalidOperationException("Failed to create workbook from template. The file is corrupted.", ce);
            else
                throw new InvalidOperationException("Failed to create workbook from template.", ce);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create workbook from template.", ex);
        }
    }

    /// <summary>
    /// 打开工作簿（简化版本）
    /// </summary>
    /// <param name="filename">文件路径</param>
    /// <param name="readOnly">是否只读</param>
    /// <param name="password">密码</param>
    /// <returns>打开的工作簿对象</returns>
    public IExcelWorkbook? OpenSimple(string filename, bool readOnly = false, string password = "")
    {
        return Open(filename, 0, readOnly, 1, password, "", false, 0, ",", true, false, 0, true);
    }

    /// <summary>
    /// 批量打开工作簿
    /// </summary>
    /// <param name="filenames">文件路径数组</param>
    /// <param name="readOnly">是否只读</param>
    /// <returns>成功打开的工作簿数量</returns>
    public int OpenRange(string[] filenames, bool readOnly = false)
    {
        if (_workbooks == null || filenames == null || filenames.Length == 0)
            return 0;

        int successCount = 0;
        foreach (string filename in filenames)
        {
            if (OpenSimple(filename, readOnly) != null)
                successCount++;
        }
        return successCount;
    }

    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据名称查找工作簿
    /// </summary>
    /// <param name="name">工作簿名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的工作簿数组</returns>
    public IExcelWorkbook[] FindByName(string name, bool matchCase = false)
    {
        if (_workbooks == null || string.IsNullOrEmpty(name) || Count == 0)
            return [];

        List<IExcelWorkbook> result = [];
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                IExcelWorkbook? workbook = this[i];
                if (workbook != null && workbook.Name != null)
                {
                    bool match = matchCase ?
                        workbook.Name.Contains(name) :
                        workbook.Name.ToLower().Contains(name.ToLower());

                    if (match)
                        result.Add(workbook);
                }
            }
            catch (Exception ex)
            {
                log.Error($"按名称查找工作簿 {name} 时，访问索引为 {i} 的工作簿发生异常", ex);
            }
        }
        return [.. result];
    }

    /// <summary>
    /// 根据路径查找工作簿
    /// </summary>
    /// <param name="path">文件路径</param>
    /// <returns>匹配的工作簿数组</returns>
    public IExcelWorkbook[] FindByPath(string path)
    {
        if (_workbooks == null || string.IsNullOrEmpty(path) || Count == 0)
            return [];

        List<IExcelWorkbook> result = [];
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                IExcelWorkbook? workbook = this[i];
                if (workbook != null && workbook.Path?.Contains(path) == true)
                {
                    result.Add(workbook);
                }
            }
            catch (Exception ex)
            {
                log.Error(ex);
            }
        }
        return [.. result];
    }


    /// <summary>
    /// 获取已保存的工作簿
    /// </summary>
    /// <returns>已保存工作簿数组</returns>
    public IExcelWorkbook[] GetSavedWorkbooks()
    {
        if (_workbooks == null || Count == 0)
            return [];

        List<IExcelWorkbook> result = [];
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                IExcelWorkbook? workbook = this[i];
                if (workbook != null && workbook.Saved)
                {
                    result.Add(workbook);
                }
            }
            catch
            {
                // 忽略单个工作簿访问异常
            }
        }
        return [.. result];
    }

    /// <summary>
    /// 获取未保存的工作簿
    /// </summary>
    /// <returns>未保存工作簿数组</returns>
    public IExcelWorkbook[] GetUnsavedWorkbooks()
    {
        if (_workbooks == null || Count == 0)
            return [];

        List<IExcelWorkbook> result = [];
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                IExcelWorkbook? workbook = this[i];
                if (workbook != null && !workbook.Saved)
                {
                    result.Add(workbook);
                }
            }
            catch (Exception x)
            {
                log.Error("获取未保存的工作簿失败：" + x.Message, x);
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取受保护的工作簿
    /// </summary>
    /// <returns>受保护工作簿数组</returns>
    public IExcelWorkbook[] GetProtectedWorkbooks()
    {
        if (_workbooks == null || Count == 0)
            return [];

        List<IExcelWorkbook> result = [];
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                IExcelWorkbook? workbook = this[i];
                if (workbook != null && workbook.IsProtected)
                {
                    result.Add(workbook);
                }
            }
            catch (Exception x)
            {
                log.Error("获取受保护的工作簿失败：" + x.Message, x);
            }
        }
        return [.. result];
    }

    /// <summary>
    /// 获取只读工作簿
    /// </summary>
    /// <returns>只读工作簿数组</returns>
    public IExcelWorkbook[] GetReadOnlyWorkbooks()
    {
        if (_workbooks == null || Count == 0)
            return [];

        var result = new System.Collections.Generic.List<IExcelWorkbook>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var workbook = this[i];
                if (workbook != null && workbook.ReadOnly)
                {
                    result.Add(workbook);
                }
            }
            catch (Exception x)
            {
                log.Error("获取只读的工作簿失败：" + x.Message, x);
            }
        }
        return [.. result];
    }

    #endregion

    #region 操作方法

    /// <summary>
    /// 关闭所有工作簿
    /// </summary>
    public void CloseAll()
    {
        try
        {
            _workbooks?.Close();
        }
        catch (Exception x)
        {
            log.Error("关闭所有工作簿时发生异常", x);
        }
    }

    /// <summary>
    /// 关闭指定索引的工作簿
    /// </summary>
    /// <param name="index">要删除的工作簿索引</param>
    /// <param name="saveChanges">是否保存更改</param>
    public void Close(int index, bool saveChanges = true)
    {
        if (_workbooks == null || index < 1 || index > Count)
            return;

        try
        {
            this[index]?.Close(saveChanges);
        }
        catch (Exception x)
        {
            log.Error($"关闭索引为 {index} 的工作簿时发生异常", x);
        }
    }

    /// <summary>
    /// 关闭指定名称的工作簿
    /// </summary>
    /// <param name="name">要关闭的工作簿名称</param>
    /// <param name="saveChanges">是否保存更改</param>
    public void Close(string name, bool saveChanges = true)
    {
        if (_workbooks == null || string.IsNullOrEmpty(name))
            return;

        try
        {
            var workbook = this[name];
            workbook?.Close(saveChanges);
        }
        catch (Exception x)
        {
            log.Error($"关闭名称为 {name} 的工作簿时发生异常", x);
        }
    }

    /// <summary>
    /// 关闭指定的工作簿
    /// </summary>
    /// <param name="workbook">要关闭的工作簿对象</param>
    /// <param name="saveChanges">是否保存更改</param>
    public void Close(IExcelWorkbook workbook, bool saveChanges = true)
    {
        if (_workbooks == null || workbook == null)
            return;

        try
        {
            workbook.Close(saveChanges);
        }
        catch (Exception x)
        {
            log.Error($"关闭工作簿时发生异常", x);
        }
    }

    /// <summary>
    /// 批量关闭工作簿
    /// </summary>
    /// <param name="names">要关闭的工作簿名称数组</param>
    /// <param name="saveChanges">是否保存更改</param>
    public void CloseRange(string[] names, bool saveChanges = true)
    {
        if (_workbooks == null || names == null || names.Length == 0)
            return;

        foreach (string name in names)
        {
            Close(name, saveChanges);
        }
    }

    /// <summary>
    /// 保存所有工作簿
    /// </summary>
    public void SaveAll()
    {
        if (_workbooks == null) return;

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                this[i]?.Save();
            }
            catch (Exception x)
            {
                log.Error($"保存索引为 {i} 的工作簿时发生异常", x);
            }
        }
    }

    /// <summary>
    /// 保存指定工作簿
    /// </summary>
    /// <param name="workbook">要保存的工作簿</param>
    public void Save(IExcelWorkbook workbook)
    {
        if (_workbooks == null || workbook == null)
            return;

        try
        {
            workbook.Save();
        }
        catch (Exception x)
        {
            log.Error($"保存工作簿时发生异常", x);
        }
    }

    /// <summary>
    /// 另存为所有工作簿
    /// </summary>
    /// <param name="folderPath">保存文件夹路径</param>
    /// <param name="fileFormat">文件格式</param>
    /// <returns>成功保存的工作簿数量</returns>
    public int SaveAsAll(string folderPath, string fileFormat = "xlsx")
    {
        if (_workbooks == null || Count == 0 || string.IsNullOrEmpty(folderPath))
            return 0;

        // 确保文件夹存在
        if (!Directory.Exists(folderPath))
            Directory.CreateDirectory(folderPath);

        int savedCount = 0;
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                if (this[i] is ExcelWorkbook workbook)
                {
                    string filename = Path.Combine(folderPath,
                        $"{workbook.Name}.{fileFormat}");
                    workbook.SaveAs(filename);
                    savedCount++;
                }
            }
            catch (Exception x)
            {
                log.Error($"保存索引为 {i} 的工作簿时发生异常", x);
            }
        }
        return savedCount;
    }

    #endregion

    #region 高级功能

    /// <summary>
    /// 获取活动工作簿
    /// </summary>
    /// <returns>活动工作簿对象</returns>
    public IExcelWorkbook? ActiveWorkbook
    {
        get
        {
            try
            {
                if (_workbooks?.Parent is not MsExcel.Application app)
                {
                    return null;
                }

                var activeWorkbook = app.ActiveWorkbook;
                return activeWorkbook != null ? new ExcelWorkbook(activeWorkbook) : null;
            }
            catch (Exception x)
            {
                log.Error("获取活动工作簿时发生异常", x);
                return null;

            }
        }
    }

    /// <summary>
    /// 获取ThisWorkbook
    /// </summary>
    /// <returns>ThisWorkbook对象</returns>
    public IExcelWorkbook? ThisWorkbook
    {
        get
        {
            try
            {
                var app = _workbooks?.Parent as MsExcel.Application;
                if (app == null)
                {
                    return null;
                }
                var thisWorkbook = app.ThisWorkbook;
                return thisWorkbook != null ? new ExcelWorkbook(thisWorkbook) : null;
            }
            catch (Exception x)
            {
                log.Error("获取ThisWorkbook时发生异常", x);
                return null;
            }
        }
    }

    /// <summary>
    /// 打印所有工作簿
    /// </summary>
    /// <param name="preview">是否打印预览</param>
    public void PrintOutAll(bool preview = false)
    {
        if (_workbooks == null) return;

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                this[i]?.PrintOut(preview);
            }
            catch (Exception x)
            {
                log.Error($"打印索引为 {i} 的工作簿时发生异常", x);
            }
        }
    }

    /// <summary>
    /// 计算所有工作簿
    /// </summary>
    public void CalculateAll()
    {
        if (_workbooks == null) return;

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                this[i]?.CalculateAll();
            }
            catch (Exception x)
            {
                log.Error($"计算索引为 {i} 的工作簿时发生异常", x);
            }
        }
    }

    /// <summary>
    /// 刷新所有工作簿
    /// </summary>
    public void RefreshAll()
    {
        if (_workbooks == null) return;

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                this[i]?.RefreshAll();
            }
            catch (Exception x)
            {
                log.Error($"刷新索引为 {i} 的工作簿时发生异常", x);
            }
        }
    }

    /// <summary>
    /// 保护所有工作簿
    /// </summary>
    /// <param name="password">保护密码</param>
    public void ProtectAll(string password = "")
    {
        if (_workbooks == null) return;

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                this[i]?.Protect(password);
            }
            catch (Exception x)
            {
                log.Error($"保护索引为 {i} 的工作簿时发生异常", x);
            }
        }
    }

    /// <summary>
    /// 取消保护所有工作簿
    /// </summary>
    /// <param name="password">保护密码</param>
    public void UnprotectAll(string password = "")
    {
        if (_workbooks == null) return;

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                this[i]?.Unprotect(password);
            }
            catch (Exception x)
            {
                log.Error($"取消保护索引为 {i} 的工作簿时发生异常", x);
            }
        }
    }

    public IEnumerator<IExcelWorkbook> GetEnumerator()
    {
        for (int i = 0; i < Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}

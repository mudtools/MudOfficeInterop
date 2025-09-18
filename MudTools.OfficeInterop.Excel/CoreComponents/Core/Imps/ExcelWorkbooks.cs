//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
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
    private MsExcel.Workbooks _workbooks;

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
    public IExcelWorkbook this[int index]
    {
        get
        {
            if (_workbooks == null || index < 1 || index > Count)
                return null;

            try
            {
                var workbook = _workbooks[index] as MsExcel.Workbook;
                return workbook != null ? new ExcelWorkbook(workbook) : null;
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
    public IExcelWorkbook this[string name]
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
    public object Parent => _workbooks?.Parent;

    /// <summary>
    /// 获取工作簿集合所在的Application对象
    /// </summary>
    public IExcelApplication Application
    {
        get
        {
            var application = _workbooks?.Application as MsExcel.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    #endregion

    #region 创建和打开

    /// <summary>
    /// 打开工作簿
    /// </summary>
    /// <param name="filename">文件路径</param>
    /// <param name="updateLinks">是否更新链接</param>
    /// <param name="readOnly">是否只读</param>
    /// <param name="format">文件格式</param>
    /// <param name="password">打开密码</param>
    /// <param name="writeResPassword">写入密码</param>
    /// <param name="ignoreReadOnlyRecommended">是否忽略只读建议</param>
    /// <param name="origin">文本来源</param>
    /// <param name="delimiter">文本分隔符</param>
    /// <param name="editable">是否可编辑</param>
    /// <param name="notify">是否通知</param>
    /// <param name="converter">格式转换器</param>
    /// <param name="addToMru">是否添加到最近使用文件</param>
    /// <returns>打开的工作簿对象</returns>
    public IExcelWorkbook Open(string filename, int updateLinks = 0, bool readOnly = false,
                              int format = 1, string password = "", string writeResPassword = "",
                              bool ignoreReadOnlyRecommended = false, int origin = 0,
                              string delimiter = ",", bool editable = true, bool notify = false,
                              int converter = 0, bool addToMru = true, object? local = null, XlCorruptLoad? corruptLoad = null)
    {
        if (_workbooks == null || string.IsNullOrEmpty(filename))
            return null;

        var corruptLoadObj = Type.Missing;
        if (corruptLoad != null)
            corruptLoadObj = (MsExcel.XlCorruptLoad)(int)corruptLoad;


        try
        {
            var workbook = _workbooks.Open(
                filename, updateLinks, readOnly, format, password, writeResPassword,
                ignoreReadOnlyRecommended, origin, delimiter, editable, notify,
                converter, addToMru, local ?? Type.Missing, corruptLoadObj);

            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
        catch
        {
            return null;
        }
    }

    public IExcelWorkbook? Add(XlWBATemplate template)
    {
        if (_workbooks == null)
            return null;
        try
        {
            MsExcel.Workbook workbook = _workbooks.Add(template);
            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
        catch
        {
            return null;
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
                workbook = _workbooks.Add() as MsExcel.Workbook;
            }
            else
            {
                workbook = _workbooks.Add(template) as MsExcel.Workbook;
            }

            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 打开工作簿（简化版本）
    /// </summary>
    /// <param name="filename">文件路径</param>
    /// <param name="readOnly">是否只读</param>
    /// <param name="password">密码</param>
    /// <returns>打开的工作簿对象</returns>
    public IExcelWorkbook OpenSimple(string filename, bool readOnly = false, string password = "")
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
                IExcelWorkbook workbook = this[i];
                if (workbook != null && workbook.Name != null)
                {
                    bool match = matchCase ?
                        workbook.Name.Contains(name) :
                        workbook.Name.ToLower().Contains(name.ToLower());

                    if (match)
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
                IExcelWorkbook workbook = this[i];
                if (workbook != null && workbook.Path?.Contains(path) == true)
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
    /// 根据修改时间查找工作簿
    /// </summary>
    /// <param name="fromTime">起始时间</param>
    /// <param name="toTime">结束时间</param>
    /// <returns>匹配的工作簿数组</returns>
    public IExcelWorkbook[] FindByModifiedTime(DateTime fromTime, DateTime toTime)
    {
        if (_workbooks == null || Count == 0)
            return new IExcelWorkbook[0];

        var result = new System.Collections.Generic.List<IExcelWorkbook>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var workbook = this[i];
                // 注意：COM对象通常不直接提供修改时间属性
                // 这里提供一个示例实现
            }
            catch
            {
                // 忽略单个工作簿访问异常
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
                IExcelWorkbook workbook = this[i];
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
                IExcelWorkbook workbook = this[i];
                if (workbook != null && !workbook.Saved)
                {
                    result.Add(workbook);
                }
            }
            catch
            {
                // 忽略单个工作簿访问异常
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
                IExcelWorkbook workbook = this[i];
                if (workbook != null && workbook.IsProtected)
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
            catch
            {
                // 忽略单个工作簿访问异常
            }
        }
        return [.. result];
    }

    #endregion

    #region 操作方法

    /// <summary>
    /// 关闭所有工作簿
    /// </summary>
    /// <param name="saveChanges">是否保存更改</param>
    public void CloseAll(bool saveChanges = true)
    {
        _workbooks?.Close();
    }

    /// <summary>
    /// 删除指定索引的工作簿
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
        catch
        {
            // 忽略关闭过程中的异常
        }
    }

    /// <summary>
    /// 删除指定名称的工作簿
    /// </summary>
    /// <param name="name">要删除的工作簿名称</param>
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
        catch
        {
            // 忽略关闭过程中的异常
        }
    }

    /// <summary>
    /// 删除指定的工作簿
    /// </summary>
    /// <param name="workbook">要删除的工作簿对象</param>
    /// <param name="saveChanges">是否保存更改</param>
    public void Close(IExcelWorkbook workbook, bool saveChanges = true)
    {
        if (_workbooks == null || workbook == null)
            return;

        try
        {
            workbook.Close(saveChanges);
        }
        catch
        {
            // 忽略关闭过程中的异常
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

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    this[i]?.Save();
                }
                catch
                {
                    // 忽略单个工作簿保存异常
                }
            }
        }
        catch
        {
            // 忽略保存过程中的异常
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
        catch
        {
            // 忽略保存过程中的异常
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

        try
        {
            // 确保文件夹存在
            if (!System.IO.Directory.Exists(folderPath))
                System.IO.Directory.CreateDirectory(folderPath);

            int savedCount = 0;
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var workbook = this[i] as ExcelWorkbook;
                    if (workbook != null)
                    {
                        string filename = System.IO.Path.Combine(folderPath,
                            $"{workbook.Name}.{fileFormat}");
                        workbook.SaveAs(filename);
                        savedCount++;
                    }
                }
                catch
                {
                    // 忽略单个工作簿另存为异常
                }
            }
            return savedCount;
        }
        catch
        {
            return 0;
        }
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
                var app = _workbooks?.Parent as MsExcel.Application;
                if (app == null)
                {
                    return null;
                }

                var activeWorkbook = app.ActiveWorkbook;
                return activeWorkbook != null ? new ExcelWorkbook(activeWorkbook) : null;
            }
            catch
            {
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
            catch
            {
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

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    this[i]?.PrintOut(preview);
                }
                catch
                {
                    // 忽略单个工作簿打印异常
                }
            }
        }
        catch
        {
            // 忽略打印过程中的异常
        }
    }

    /// <summary>
    /// 计算所有工作簿
    /// </summary>
    public void CalculateAll()
    {
        if (_workbooks == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    this[i]?.CalculateAll();
                }
                catch
                {
                    // 忽略单个工作簿计算异常
                }
            }
        }
        catch
        {
            // 忽略计算过程中的异常
        }
    }

    /// <summary>
    /// 刷新所有工作簿
    /// </summary>
    public void RefreshAll()
    {
        if (_workbooks == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    this[i]?.RefreshAll();
                }
                catch
                {
                    // 忽略单个工作簿刷新异常
                }
            }
        }
        catch
        {
            // 忽略刷新过程中的异常
        }
    }

    /// <summary>
    /// 保护所有工作簿
    /// </summary>
    /// <param name="password">保护密码</param>
    public void ProtectAll(string password = "")
    {
        if (_workbooks == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    this[i]?.Protect(password);
                }
                catch
                {
                    // 忽略单个工作簿保护异常
                }
            }
        }
        catch
        {
            // 忽略保护过程中的异常
        }
    }

    /// <summary>
    /// 取消保护所有工作簿
    /// </summary>
    /// <param name="password">保护密码</param>
    public void UnprotectAll(string password = "")
    {
        if (_workbooks == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    this[i]?.Unprotect(password);
                }
                catch
                {
                    // 忽略单个工作簿取消保护异常
                }
            }
        }
        catch
        {
            // 忽略取消保护过程中的异常
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

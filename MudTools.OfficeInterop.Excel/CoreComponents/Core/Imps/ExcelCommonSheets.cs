//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel工作表集合的公共基类实现
/// </summary>
internal abstract class ExcelCommonSheets : IExcelCommonSheets
{
    #region IDisposable Support
    protected bool _disposedValue = false;

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                Marshal.ReleaseComObject(_hpageBreaks);
                Marshal.ReleaseComObject(_vpageBreaks);
            }
            _hpageBreaks = null;
            _vpageBreaks = null;
            _disposedValue = true;
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
    #endregion

    #region IEnumerable<IExcelWorksheet> Support
    public abstract IEnumerator<IExcelWorksheet> GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
    #endregion

    #region IExcelCommonSheets Implementation

    /// <summary>
    /// 获取工作表集合中的工作表数量
    /// </summary>
    public abstract int Count { get; }

    /// <summary>
    /// 获取指定索引的工作表对象
    /// </summary>
    /// <param name="index">工作表索引（从1开始）</param>
    /// <returns>工作表对象</returns>
    public abstract IExcelWorksheet this[int index] { get; }

    /// <summary>
    /// 获取指定名称的工作表对象
    /// </summary>
    /// <param name="name">工作表名称</param>
    /// <returns>工作表对象</returns>
    public IExcelWorksheet this[string name]
    {
        get
        {
            IExcelWorksheet[] result = FindByName(name);
            if (result != null && result.Length > 0)
                return result[0];
            return null;
        }
    }

    public IExcelWorksheet[] this[params string[] names]
    {
        get
        {
            if (names == null)
                return [];
            if (names.Length < 1)
                return [];

            List<IExcelWorksheet> results = [];
            foreach (string name in names)
            {
                IExcelWorksheet[] result = FindByName(name);
                if (result != null && result.Length > 0)
                    results.AddRange(result);
            }
            return [.. results];
        }
    }

    protected abstract object NativeSheets { get; }

    /// <summary>
    /// 获取工作表集合所在的父对象（通常是 Workbook）
    /// </summary>
    public abstract object Parent { get; }

    /// <summary>
    /// 获取工作表集合所在的Application对象
    /// </summary>
    public abstract IExcelApplication Application { get; }


    private IExcelHPageBreaks? _hpageBreaks = null;

    public IExcelHPageBreaks? HPageBreaks
    {
        get
        {
            if (NativeSheets == null)
                return null;
            if (_hpageBreaks != null)
                return _hpageBreaks;
            if (NativeSheets is MsExcel.Sheets sheets)
            {
                _hpageBreaks = new ExcelHPageBreaks(sheets.HPageBreaks);
            }
            else if (NativeSheets is MsExcel.Worksheets wsheets)
            {
                _hpageBreaks = new ExcelHPageBreaks(wsheets.HPageBreaks);
            }
            return _hpageBreaks;
        }
    }

    private IExcelVPageBreaks? _vpageBreaks = null;

    public IExcelVPageBreaks? VPageBreaks
    {
        get
        {
            if (NativeSheets == null)
                return null;
            if (_vpageBreaks != null)
                return _vpageBreaks;
            if (NativeSheets is MsExcel.Sheets sheets)
            {
                _vpageBreaks = new ExcelVPageBreaks(sheets.VPageBreaks);
            }
            else if (NativeSheets is MsExcel.Worksheets wsheets)
            {
                _vpageBreaks = new ExcelVPageBreaks(wsheets.VPageBreaks);
            }
            return _vpageBreaks;
        }
    }

    /// <summary>
    /// 根据名称查找工作表
    /// </summary>
    /// <param name="name">工作表名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的工作表数组</returns>
    public virtual IExcelWorksheet[] FindByName(string name, bool matchCase = false)
    {
        if (string.IsNullOrEmpty(name) || Count == 0)
            return [];

        List<IExcelWorksheet> result = [];
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                IExcelWorksheet worksheet = this[i];
                if (worksheet != null && worksheet.Name != null)
                {
                    bool match = matchCase ?
                        worksheet.Name.Contains(name) :
                        worksheet.Name.ToLower().Contains(name.ToLower());

                    if (match)
                        result.Add(worksheet);
                }
            }
            catch
            {
                // 忽略单个工作表访问异常
            }
        }
        return [.. result];
    }

    /// <summary>
    /// 根据类型查找工作表 (例如: xlWorksheet, xlChart, xlExcel4MacroSheet, xlExcel4IntlMacroSheet)
    /// </summary>
    /// <param name="type">工作表类型</param>
    /// <returns>匹配的工作表数组</returns>
    public virtual IExcelWorksheet[] FindByType(XlSheetType type)
    {
        if (Count == 0)
            return [];

        List<IExcelWorksheet> result = [];
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                IExcelWorksheet worksheet = this[i];
                if (worksheet != null && worksheet.Type == type)
                {
                    result.Add(worksheet);
                }
            }
            catch
            {
                // 忽略单个工作表访问异常
            }
        }
        return [.. result];
    }

    /// <summary>
    /// 根据索引范围查找工作表
    /// </summary>
    /// <param name="startIndex">起始索引</param>
    /// <param name="endIndex">结束索引</param>
    /// <returns>匹配的工作表数组</returns>
    public virtual IExcelWorksheet[] FindByIndexRange(int startIndex, int endIndex)
    {
        if (Count == 0 || startIndex < 1 || endIndex > Count)
            return [];

        List<IExcelWorksheet> result = [];
        for (int i = startIndex; i <= Math.Min(endIndex, Count); i++)
        {
            IExcelWorksheet worksheet = this[i];
            if (worksheet != null)
                result.Add(worksheet);
        }
        return [.. result];
    }

    /// <summary>
    /// 删除指定索引的工作表
    /// </summary>
    /// <param name="index">要删除的工作表索引</param>
    public abstract void Delete(int index);

    /// <summary>
    /// 删除指定名称的工作表
    /// </summary>
    /// <param name="name">要删除的工作表名称</param>
    public abstract void Delete(string name);

    /// <summary>
    /// 删除指定的工作表对象
    /// </summary>
    /// <param name="sheet">要删除的工作表对象</param>
    public abstract void Delete(IExcelWorksheet sheet);


    /// <summary>
    /// 添加新工作表
    /// </summary>
    /// <param name="options">添加选项</param>
    /// <returns>新创建的工作表</returns>
    public IExcelWorksheet AddSheet(AddSheetOptions options)
    {
        if (options == null)
            options = new AddSheetOptions();

        // 验证参数
        if (options.Before != null && options.After != null)
            throw new ArgumentException("不能同时指定Before和After参数");

        if (options.Count < 1 || options.Count > 50)
            throw new ArgumentOutOfRangeException(nameof(options.Count), "添加数量必须在1-50之间");

        try
        {
            // 准备参数
            object beforeObj = options.Before != null ?
                ((ExcelWorksheet)options.Before).Worksheet :
                Type.Missing;

            object afterObj = options.After != null ?
                ((ExcelWorksheet)options.After).Worksheet :
                Type.Missing;

            object countObj = options.Count;
            object typeObj = (MsExcel.XlSheetType)options.Type;

            // 处理模板
            object templateObj = options.Template != null ?
                ((ExcelWorksheet)options.Template).Worksheet :
                Type.Missing;
            MsExcel.Worksheet? newSheet = null;
            if (NativeSheets is MsExcel.Sheets _nativeSheets)
            {
                // 添加新工作表
                newSheet = _nativeSheets.Add(
                    Before: beforeObj,
                    After: afterObj,
                    Count: countObj,
                    Type: typeObj
                ) as MsExcel.Worksheet;
            }
            else if (NativeSheets is MsExcel.Worksheets _nativeWorkSheets)
            {
                // 添加新工作表
                newSheet = _nativeWorkSheets.Add(
                    Before: beforeObj,
                    After: afterObj,
                    Count: countObj,
                    Type: typeObj
                ) as MsExcel.Worksheet;
            }

            if (newSheet == null)
                throw new InvalidOperationException("添加工作表失败");

            var result = new ExcelWorksheet(newSheet);

            // 设置名称（自动处理重复）
            if (!string.IsNullOrEmpty(options.Name))
            {
                try
                {
                    result.Name = options.Name;
                }
                catch (COMException) // 名称重复错误
                {
                    // 自动生成唯一名称
                    int index = 1;
                    string baseName = options.Name.Trim();
                    while (index < 100)
                    {
                        try
                        {
                            result.Name = $"{baseName}_{index}";
                            break;
                        }
                        catch
                        {
                            index++;
                        }
                    }
                }
            }

            // 应用模板格式
            //if (options.Template != null)
            //{
            //    result.UsedRange.CopyFrom(options.Template.UsedRange);
            //}

            return result;
        }
        catch (COMException ex)
        {
            throw new ExcelOperationException("添加工作表失败", ex);
        }
    }

    /// <summary>
    /// 复制工作表
    /// </summary>
    /// <param name="source">源工作表</param>
    /// <param name="options">复制选项</param>
    /// <returns>新创建的工作表副本</returns>
    public IExcelWorksheet CopySheet(IExcelWorksheet source, CopySheetOptions options)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));

        if (options == null)
            options = new CopySheetOptions();

        // 验证参数
        if (options.Before != null && options.After != null)
            throw new ArgumentException("不能同时指定Before和After参数");

        try
        {
            // 获取底层工作表对象
            var nativeSheet = (source as ExcelWorksheet)?.Worksheet;
            if (nativeSheet == null)
                throw new ArgumentException("无效的工作表对象", nameof(source));

            // 处理跨工作簿复制
            MsExcel.Workbook targetWorkbook = null;
            if (options.TargetWorkbook != null)
            {
                targetWorkbook = (options.TargetWorkbook as ExcelWorkbook)?._workbook;
                if (targetWorkbook == null)
                    throw new ArgumentException("无效的目标工作簿", nameof(options.TargetWorkbook));
            }

            // 准备复制参数
            object beforeObj = options.Before != null ?
                ((ExcelWorksheet)options.Before).Worksheet :
                Type.Missing;

            object afterObj = options.After != null ?
                ((ExcelWorksheet)options.After).Worksheet :
                Type.Missing;

            // 执行复制
            nativeSheet.Copy(beforeObj, afterObj);

            // 获取新创建的工作表（总是位于复制位置）
            MsExcel.Worksheet? newSheet = null;
            if (targetWorkbook != null)
            {
                // 跨工作簿复制时，新工作表在目标工作簿的活动表位置
                newSheet = targetWorkbook.ActiveSheet as MsExcel.Worksheet;
            }
            else
            {
                // 同一工作簿内复制
                if (NativeSheets is MsExcel.Sheets _nativeSheets)
                {
                    // 同一工作簿内复制
                    int newIndex = options.Before != null ?
                    options.Before.Index :
                    (options.After?.Index ?? _nativeSheets.Count) + 1;

                    newSheet = _nativeSheets[newIndex] as MsExcel.Worksheet;
                }
                else if (NativeSheets is MsExcel.Worksheets _nativeworkSheets)
                {
                    // 同一工作簿内复制
                    int newIndex = options.Before != null ?
                    options.Before.Index :
                    (options.After?.Index ?? _nativeworkSheets.Count) + 1;
                    newSheet = _nativeworkSheets[newIndex] as MsExcel.Worksheet;
                }
            }
            if (newSheet == null)
            {
                throw new InvalidOperationException("工作表复制失败");
            }

            var result = new ExcelWorksheet(newSheet);

            // 处理仅复制值的情况
            if (options.ValuesOnly)
            {
                result.UsedRange.Value = source.UsedRange.Value;
            }

            return result;
        }
        catch (COMException ex)
        {
            throw new ExcelOperationException("工作表复制失败", ex);
        }
    }

    /// <summary>
    /// 批量删除工作表
    /// </summary>
    /// <param name="indices">要删除的工作表索引数组</param>
    public virtual void DeleteRange(int[] indices)
    {
        if (indices == null || indices.Length == 0)
            return;

        // 按索引从大到小排序，避免删除过程中索引变化的问题
        List<int> sortedIndices = indices.OrderByDescending(x => x).ToList();
        foreach (int index in sortedIndices)
        {
            Delete(index);
        }
    }

    /// <summary>
    /// 批量删除工作表
    /// </summary>
    /// <param name="names">要删除的工作表名称数组</param>
    public virtual void DeleteRange(string[] names)
    {
        if (names == null || names.Length == 0)
            return;

        foreach (string name in names)
        {
            Delete(name);
        }
    }

    /// <summary>
    /// 选择多个工作表
    /// </summary>
    /// <param name="worksheetNames">工作表名称数组</param>
    public abstract void Select(params string[] worksheetNames);

    /// <summary>
    /// 获取活动工作表
    /// </summary>
    /// <returns>活动工作表对象</returns>
    public abstract IExcelWorksheet ActiveWorksheet { get; }

    /// <summary>
    /// 打印所有工作表
    /// </summary>
    /// <param name="preview">是否打印预览</param>
    public abstract void PrintOutAll(bool preview = false);

    /// <summary>
    /// 计算所有工作表
    /// </summary>
    public abstract void Calculate();

    /// <summary>
    /// 刷新所有工作表
    /// </summary>
    public abstract void RefreshAll();

    /// <summary>
    /// 保护所有工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    public virtual void ProtectAll(string password = "")
    {
        if (Count == 0) return;

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
                    // 忽略单个工作表保护异常
                }
            }
        }
        catch
        {
            // 忽略保护过程中的异常
        }
    }

    /// <summary>
    /// 取消保护所有工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    public virtual void UnprotectAll(string password = "")
    {
        if (Count == 0) return;

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
                    // 忽略单个工作表取消保护异常
                }
            }
        }
        catch
        {
            // 忽略取消保护过程中的异常
        }
    }

    #endregion
}
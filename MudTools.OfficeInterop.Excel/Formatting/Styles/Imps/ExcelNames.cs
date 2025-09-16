//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


using log4net;

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel Names 集合对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Names 对象的安全访问和资源管理
/// </summary>
internal class ExcelNames : IExcelNames
{
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelNames));
    /// <summary>
    /// 底层的 COM Names 集合对象
    /// </summary>
    private MsExcel.Names _names;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelNames 实例
    /// </summary>
    /// <param name="names">底层的 COM Names 集合对象</param>
    internal ExcelNames(MsExcel.Names names)
    {
        _names = names;
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
            try
            {
                // 释放所有子名称对象
                for (int i = 1; i <= Count; i++)
                {
                    var name = this[i] as ExcelName;
                    name?.Dispose();
                }

                // 释放底层COM对象
                if (_names != null)
                    Marshal.ReleaseComObject(_names);
            }
            catch (Exception ex)
            {
                log.Warn("释放ExcelNames资源时发生异常", ex);
                // 忽略释放过程中的异常
            }
            _names = null;
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
    /// 获取名称集合中的名称数量
    /// </summary>
    public int Count => _names?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的名称对象
    /// </summary>
    /// <param name="index">名称索引（从1开始）</param>
    /// <returns>名称对象</returns>
    public IExcelName this[int index]
    {
        get
        {
            if (_names == null || index < 1 || index > Count)
                return null;

            try
            {
                var name = _names.Item(index);
                return name != null ? new ExcelName(name) : null;
            }
            catch (Exception ex)
            {
                log.Warn($"获取索引为 {index} 的名称时发生异常", ex);
                return null;
            }
        }
    }

    /// <summary>
    /// 获取指定名称的名称对象
    /// </summary>
    /// <param name="name">名称</param>
    /// <returns>名称对象</returns>
    public IExcelName this[string name]
    {
        get
        {
            if (_names == null || string.IsNullOrEmpty(name))
                return null;

            try
            {
                var excelName = _names.Item(name);
                return excelName != null ? new ExcelName(excelName) : null;
            }
            catch (Exception ex)
            {
                log.Warn($"获取名称为 '{name}' 的名称对象时发生异常", ex);
                return null;
            }
        }
    }

    /// <summary>
    /// 获取名称集合所在的父对象
    /// </summary>
    public object Parent => _names?.Parent;

    /// <summary>
    /// 获取名称集合所在的Application对象
    /// </summary>
    public IExcelApplication Application
    {
        get
        {
            var application = _names?.Application as MsExcel.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    #endregion

    #region 创建和添加

    /// <summary>
    /// 添加新的名称
    /// </summary>
    /// <param name="name">名称</param>
    /// <param name="refersTo">引用</param>
    /// <param name="visible">是否可见</param>
    /// <param name="macroType">宏类型</param>
    /// <param name="shortcutKey">快捷键</param>
    /// <param name="category">类别</param>
    /// <param name="nameLocal">本地名称</param>
    /// <param name="refersToLocal">本地引用</param>
    /// <param name="categoryLocal">本地类别</param>
    /// <param name="refersToR1C1">R1C1引用</param>
    /// <param name="refersToR1C1Local">本地R1C1引用</param>
    /// <returns>新创建的名称对象</returns>
    public IExcelName? Add(string name, object? refersTo = null, bool visible = true,
                         int macroType = 0, string shortcutKey = "", object? category = null,
                         string nameLocal = "", object? refersToLocal = null, object? categoryLocal = null,
                         string refersToR1C1 = "", string refersToR1C1Local = "")
    {
        if (_names == null || string.IsNullOrEmpty(name))
            return null;

        try
        {
            var excelName = _names.Add(
                name, refersTo, visible, macroType, shortcutKey, category,
                nameLocal, refersToLocal, categoryLocal, refersToR1C1, refersToR1C1Local
            ) as MsExcel.Name;

            return excelName != null ? new ExcelName(excelName) : null;
        }
        catch (Exception ex)
        {
            log.Error($"添加名称 '{name}' 时发生异常", ex);
            return null;
        }
    }

    /// <summary>
    /// 基于区域创建名称
    /// </summary>
    /// <param name="range">区域对象</param>
    /// <param name="name">名称</param>
    /// <param name="useColumnNames">是否使用列名</param>
    /// <param name="useRowNames">是否使用行名</param>
    /// <returns>创建的名称对象</returns>
    public IExcelName? CreateFromRange(IExcelRange range, string name = "",
                                    bool useColumnNames = false, bool useRowNames = false)
    {
        if (_names == null || range == null)
            return null;

        try
        {
            // 通过区域创建名称的逻辑实现
            var excelRange = range as ExcelRange;
            if (excelRange?.InternalRange != null)
            {
                string rangeName = !string.IsNullOrEmpty(name) ? name : $"Range_{DateTime.Now:yyyyMMddHHmmss}";
                string refersTo = excelRange.InternalRange.Address;

                return Add(rangeName, refersTo);
            }
            return null;
        }
        catch (Exception ex)
        {
            log.Error("基于区域创建名称时发生异常", ex);
            return null;
        }
    }

    /// <summary>
    /// 创建工作表名称
    /// </summary>
    /// <param name="worksheet">工作表对象</param>
    /// <param name="name">名称</param>
    /// <returns>创建的名称对象</returns>
    public IExcelName? CreateWorksheetName(IExcelWorksheet worksheet, string name = "")
    {
        if (_names == null || worksheet == null)
            return null;

        try
        {
            // 通过工作表创建名称的逻辑实现
            string worksheetName = !string.IsNullOrEmpty(name) ? name : worksheet.Name;
            string refersTo = worksheet.Name;

            return Add(worksheetName, refersTo);
        }
        catch (Exception ex)
        {
            log.Error("创建工作表名称时发生异常", ex);
            return null;
        }
    }

    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据名称查找
    /// </summary>
    /// <param name="name">名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的名称数组</returns>
    public IExcelName[] FindByName(string name, bool matchCase = false)
    {
        if (_names == null || string.IsNullOrEmpty(name) || Count == 0)
            return new IExcelName[0];

        var result = new System.Collections.Generic.List<IExcelName>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var excelName = this[i];
                if (excelName != null && excelName.Name != null)
                {
                    bool match = matchCase ?
                        excelName.Name.Contains(name) :
                        excelName.Name.ToLower().Contains(name.ToLower());

                    if (match)
                        result.Add(excelName);
                }
            }
            catch (Exception ex)
            {
                log.Warn($"查找名称过程中访问索引为 {i} 的名称时发生异常", ex);
                // 忽略单个名称访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据引用查找
    /// </summary>
    /// <param name="refersTo">引用</param>
    /// <returns>匹配的名称数组</returns>
    public IExcelName[] FindByRefersTo(string refersTo)
    {
        if (_names == null || string.IsNullOrEmpty(refersTo) || Count == 0)
            return new IExcelName[0];

        var result = new System.Collections.Generic.List<IExcelName>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var excelName = this[i];
                if (excelName != null && excelName.RefersTo?.Contains(refersTo) == true)
                {
                    result.Add(excelName);
                }
            }
            catch (Exception ex)
            {
                log.Warn($"查找引用过程中访问索引为 {i} 的名称时发生异常", ex);
                // 忽略单个名称访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据可见性查找
    /// </summary>
    /// <param name="visible">可见性</param>
    /// <returns>匹配的名称数组</returns>
    public IExcelName[] FindByVisibility(bool visible)
    {
        if (_names == null || Count == 0)
            return new IExcelName[0];

        var result = new System.Collections.Generic.List<IExcelName>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var excelName = this[i];
                if (excelName != null && excelName.Visible == visible)
                {
                    result.Add(excelName);
                }
            }
            catch (Exception ex)
            {
                log.Warn($"根据可见性查找过程中访问索引为 {i} 的名称时发生异常", ex);
                // 忽略单个名称访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据类别查找
    /// </summary>
    /// <param name="category">类别</param>
    /// <returns>匹配的名称数组</returns>
    public IExcelName[] FindByCategory(string category)
    {
        if (_names == null || string.IsNullOrEmpty(category) || Count == 0)
            return new IExcelName[0];

        var result = new System.Collections.Generic.List<IExcelName>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var excelName = this[i];
                if (excelName != null && excelName.Category?.ToString()?.Contains(category) == true)
                {
                    result.Add(excelName);
                }
            }
            catch (Exception ex)
            {
                log.Warn($"根据类别查找过程中访问索引为 {i} 的名称时发生异常", ex);
                // 忽略单个名称访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取可见的名称
    /// </summary>
    /// <returns>可见名称数组</returns>
    public IExcelName[] GetVisibleNames()
    {
        return FindByVisibility(true);
    }

    /// <summary>
    /// 获取隐藏的名称
    /// </summary>
    /// <returns>隐藏名称数组</returns>
    public IExcelName[] GetHiddenNames()
    {
        return FindByVisibility(false);
    }

    /// <summary>
    /// 获取工作簿级别的名称
    /// </summary>
    /// <returns>工作簿级别名称数组</returns>
    public IExcelName[] GetWorkbookNames()
    {
        if (_names == null || Count == 0)
            return new IExcelName[0];

        var result = new System.Collections.Generic.List<IExcelName>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var excelName = this[i];
                // 检查名称是否为工作簿级别（父对象为工作簿）
                if (excelName != null && excelName.Parent is MsExcel.Workbook)
                {
                    result.Add(excelName);
                }
            }
            catch (Exception ex)
            {
                log.Warn($"获取工作簿级别名称过程中访问索引为 {i} 的名称时发生异常", ex);
                // 忽略单个名称访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取工作表级别的名称
    /// </summary>
    /// <returns>工作表级别名称数组</returns>
    public IExcelName[] GetWorksheetNames()
    {
        if (_names == null || Count == 0)
            return new IExcelName[0];

        var result = new System.Collections.Generic.List<IExcelName>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var excelName = this[i];
                // 检查名称是否为工作表级别（父对象为工作表）
                if (excelName != null && excelName.Parent is MsExcel.Worksheet)
                {
                    result.Add(excelName);
                }
            }
            catch (Exception ex)
            {
                log.Warn($"获取工作表级别名称过程中访问索引为 {i} 的名称时发生异常", ex);
                // 忽略单个名称访问异常
            }
        }
        return result.ToArray();
    }

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除所有名称
    /// </summary>
    public void Clear()
    {
        if (_names == null) return;

        try
        {
            // 从后往前删除，避免索引变化问题
            for (int i = Count; i >= 1; i--)
            {
                try
                {
                    _names.Item(i).Delete();
                }
                catch (Exception ex)
                {
                    log.Warn($"清空名称时删除索引为 {i} 的名称时发生异常", ex);
                    // 忽略删除过程中的异常
                }
            }
        }
        catch (Exception ex)
        {
            log.Warn("清空所有名称时发生异常", ex);
            // 忽略清空过程中的异常
        }
    }

    /// <summary>
    /// 删除指定索引的名称
    /// </summary>
    /// <param name="index">要删除的名称索引</param>
    public void Delete(int index)
    {
        if (_names == null || index < 1 || index > Count)
            return;

        try
        {
            _names.Item(index).Delete();
        }
        catch (Exception ex)
        {
            log.Warn($"删除索引为 {index} 的名称时发生异常", ex);
            // 忽略删除过程中的异常
        }
    }

    /// <summary>
    /// 删除指定名称的名称
    /// </summary>
    /// <param name="name">要删除的名称</param>
    public void Delete(string name)
    {
        if (_names == null || string.IsNullOrEmpty(name))
            return;

        try
        {
            var excelName = _names.Item(name);
            excelName?.Delete();
        }
        catch (Exception ex)
        {
            log.Warn($"删除名称 '{name}' 时发生异常", ex);
            // 忽略删除过程中的异常
        }
    }

    /// <summary>
    /// 删除指定的名称对象
    /// </summary>
    /// <param name="nameObject">要删除的名称对象</param>
    public void Delete(IExcelName nameObject)
    {
        if (_names == null || nameObject == null)
            return;

        try
        {
            nameObject.Delete();
        }
        catch (Exception ex)
        {
            log.Warn("删除指定名称对象时发生异常", ex);
            // 忽略删除过程中的异常
        }
    }

    /// <summary>
    /// 批量删除名称
    /// </summary>
    /// <param name="names">要删除的名称数组</param>
    public void DeleteRange(string[] names)
    {
        if (_names == null || names == null || names.Length == 0)
            return;

        foreach (string name in names)
        {
            Delete(name);
        }
    }

    /// <summary>
    /// 选择所有名称
    /// </summary>
    public void SelectAll()
    {
        // Excel中没有直接选择所有名称的方法
        // 这里提供一个空实现以保持接口一致性
    }

    /// <summary>
    /// 取消选择所有名称
    /// </summary>
    public void DeselectAll()
    {
        // Excel中没有直接取消选择的方法
        // 这里提供一个空实现以保持接口一致性
    }

    /// <summary>
    /// 刷新所有名称
    /// </summary>
    public void Refresh()
    {
        // Excel名称通常会自动刷新
        // 这里提供一个空实现以保持接口一致性
    }

    #endregion

    #region 导出和导入

    /// <summary>
    /// 导出所有名称到文本文件
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    /// <param name="includeHidden">是否包含隐藏名称</param>
    /// <returns>是否导出成功</returns>
    public bool ExportToText(string filename, bool includeHidden = false)
    {
        if (_names == null || Count == 0 || string.IsNullOrEmpty(filename))
            return false;

        try
        {
            using (var writer = new System.IO.StreamWriter(filename, false, System.Text.Encoding.UTF8))
            {
                writer.WriteLine("Excel Names Export");
                writer.WriteLine("==================");
                writer.WriteLine($"Export Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                writer.WriteLine($"Total Names: {Count}");
                writer.WriteLine($"Include Hidden: {includeHidden}");
                writer.WriteLine();

                for (int i = 1; i <= Count; i++)
                {
                    try
                    {
                        var excelName = this[i];
                        if (excelName != null && (excelName.Visible || includeHidden))
                        {
                            writer.WriteLine($"Name #{i}");
                            writer.WriteLine($"Name: {excelName.Name}");
                            writer.WriteLine($"NameLocal: {excelName.NameLocal}");
                            writer.WriteLine($"RefersTo: {excelName.RefersTo}");
                            writer.WriteLine($"RefersToLocal: {excelName.RefersToLocal}");
                            writer.WriteLine($"RefersToR1C1: {excelName.RefersToR1C1}");
                            writer.WriteLine($"RefersToR1C1Local: {excelName.RefersToR1C1Local}");
                            writer.WriteLine($"Visible: {excelName.Visible}");
                            writer.WriteLine($"Category: {excelName.Category}");
                            writer.WriteLine($"MacroType: {excelName.MacroType}");
                            writer.WriteLine($"ShortcutKey: {excelName.ShortcutKey}");
                            writer.WriteLine($"Comment: {excelName.Comment}");
                            writer.WriteLine(new string('-', 50));
                            writer.WriteLine();
                        }
                    }
                    catch (Exception ex)
                    {
                        log.Warn($"导出名称过程中访问索引为 {i} 的名称时发生异常", ex);
                        // 忽略单个名称导出异常
                    }
                }
            }
            return true;
        }
        catch (Exception ex)
        {
            log.Error($"导出名称到文件 '{filename}' 时发生异常", ex);
            return false;
        }
    }

    /// <summary>
    /// 从文本文件导入名称
    /// </summary>
    /// <param name="filename">导入文件路径</param>
    /// <returns>成功导入的名称数量</returns>
    public int ImportFromText(string filename)
    {
        if (_names == null || string.IsNullOrEmpty(filename))
            return 0;

        // 注意：Excel Names不支持直接导入
        // 这里提供一个示例实现框架
        return 0;
    }
    #endregion

    #region 高级功能

    /// <summary>
    /// 获取活动名称
    /// </summary>
    /// <returns>活动名称对象</returns>
    public IExcelName ActiveName
    {
        get
        {
            // Excel中没有直接的活动名称概念
            // 这里提供一个空实现以保持接口一致性
            return null;
        }
    }
    #endregion

    #region 私有辅助方法

    /// <summary>
    /// 获取引用类型
    /// </summary>
    /// <param name="reference">引用</param>
    /// <returns>引用类型</returns>
    private string GetReferenceType(string reference)
    {
        if (string.IsNullOrEmpty(reference))
            return "Empty";

        if (reference.StartsWith("="))
            return "Formula";

        if (reference.StartsWith("#"))
            return "Error";

        if (reference.All(c => char.IsDigit(c) || c == '.' || c == '-'))
            return "Constant";

        if (reference.Contains(":") || reference.Contains("$") ||
            reference.Any(c => char.IsLetter(c)))
            return "Range";

        return "Other";
    }

    /// <summary>
    /// 从引用中获取工作表名称
    /// </summary>
    /// <param name="reference">引用</param>
    /// <returns>工作表名称</returns>
    private string GetWorksheetNameFromReference(string reference)
    {
        if (string.IsNullOrEmpty(reference) || !reference.Contains("!"))
            return "";

        try
        {
            string[] parts = reference.Split('!');
            return parts.Length > 0 ? parts[0].Trim('\'', '"') : "";
        }
        catch (Exception ex)
        {
            log.Warn("从引用中获取工作表名称时发生异常", ex);
            return "";
        }
    }

    /// <summary>
    /// 从引用中获取区域地址
    /// </summary>
    /// <param name="reference">引用</param>
    /// <returns>区域地址</returns>
    private string GetRangeAddressFromReference(string reference)
    {
        if (string.IsNullOrEmpty(reference))
            return "";

        try
        {
            if (reference.Contains("!"))
            {
                string[] parts = reference.Split('!');
                return parts.Length > 1 ? parts[1] : reference;
            }
            return reference;
        }
        catch (Exception ex)
        {
            log.Warn("从引用中获取区域地址时发生异常", ex);
            return "";
        }
    }

    /// <summary>
    /// 判断是否为区域引用
    /// </summary>
    /// <param name="reference">引用</param>
    /// <returns>是否为区域引用</returns>
    private bool IsRangeReference(string reference)
    {
        if (string.IsNullOrEmpty(reference))
            return false;

        return reference.Contains(":") || reference.Contains("$") ||
               reference.Any(c => char.IsLetter(c));
    }

    /// <summary>
    /// 判断是否为常量引用
    /// </summary>
    /// <param name="reference">引用</param>
    /// <returns>是否为常量引用</returns>
    private bool IsConstantReference(string reference)
    {
        if (string.IsNullOrEmpty(reference))
            return false;

        return reference.All(c => char.IsDigit(c) || c == '.' || c == '-');
    }

    /// <summary>
    /// 判断是否为公式引用
    /// </summary>
    /// <param name="reference">引用</param>
    /// <returns>是否为公式引用</returns>
    private bool IsFormulaReference(string reference)
    {
        if (string.IsNullOrEmpty(reference))
            return false;

        return reference.StartsWith("=");
    }

    public IEnumerator<IExcelName> GetEnumerator()
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
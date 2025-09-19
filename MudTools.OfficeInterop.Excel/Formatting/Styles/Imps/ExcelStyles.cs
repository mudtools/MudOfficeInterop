//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


using System.Drawing;

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel Styles 集合对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Styles 对象的安全访问和资源管理
/// </summary>
internal class ExcelStyles : IExcelStyles
{
    /// <summary>
    /// 底层的 COM Styles 集合对象
    /// </summary>
    private MsExcel.Styles _styles;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelStyles 实例
    /// </summary>
    /// <param name="styles">底层的 COM Styles 集合对象</param>
    internal ExcelStyles(MsExcel.Styles styles)
    {
        _styles = styles;
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
            if (_styles != null)
                Marshal.ReleaseComObject(_styles);
            _styles = null;
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
    /// 获取样式集合中的样式数量
    /// </summary>
    public int Count => _styles?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的样式对象
    /// </summary>
    /// <param name="index">样式索引（从1开始）</param>
    /// <returns>样式对象</returns>
    public IExcelStyle this[int index]
    {
        get
        {
            if (_styles == null || index < 1 || index > Count)
                return null;

            try
            {
                var style = _styles[index] as MsExcel.Style;
                return style != null ? new ExcelStyle(style) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取指定名称的样式对象
    /// </summary>
    /// <param name="name">样式名称</param>
    /// <returns>样式对象</returns>
    public IExcelStyle this[string name]
    {
        get
        {
            if (_styles == null || string.IsNullOrEmpty(name))
                return null;

            try
            {
                var style = _styles[name] as MsExcel.Style;
                return style != null ? new ExcelStyle(style) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取样式集合所在的父对象
    /// </summary>
    public object Parent => _styles?.Parent;

    /// <summary>
    /// 获取样式集合所在的Application对象
    /// </summary>
    public IExcelApplication Application
    {
        get
        {
            var application = _styles?.Application as MsExcel.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    #endregion

    #region 创建和添加

    /// <summary>
    /// 添加新的样式
    /// </summary>
    /// <param name="name">样式名称</param>
    /// <returns>新创建的样式对象</returns>
    public IExcelStyle Add(string name)
    {
        if (_styles == null || string.IsNullOrEmpty(name))
            return null;

        try
        {
            var style = _styles.Add(name) as MsExcel.Style;
            return style != null ? new ExcelStyle(style) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 基于现有样式创建新样式
    /// </summary>
    /// <param name="name">新样式名称</param>
    /// <param name="basedOn">基础样式</param>
    /// <returns>新创建的样式对象</returns>
    public IExcelStyle AddBasedOn(string name, IExcelStyle basedOn)
    {
        if (_styles == null || string.IsNullOrEmpty(name) || basedOn == null)
            return null;

        try
        {
            // 先添加新样式
            var newStyle = Add(name);
            if (newStyle != null)
            {
                // 复制基础样式的属性
                var excelBasedOn = basedOn as ExcelStyle;
                if (excelBasedOn?._style != null)
                {
                    // 复制各种样式属性
                    newStyle.Font.Name = excelBasedOn.Font.Name;
                    newStyle.Font.Size = excelBasedOn.Font.Size;
                    newStyle.Font.Bold = excelBasedOn.Font.Bold;
                    newStyle.Font.Italic = excelBasedOn.Font.Italic;
                    newStyle.Font.Color = excelBasedOn.Font.Color;
                    newStyle.NumberFormat = excelBasedOn.NumberFormat;
                    newStyle.HorizontalAlignment = excelBasedOn.HorizontalAlignment;
                    newStyle.VerticalAlignment = excelBasedOn.VerticalAlignment;
                    newStyle.WrapText = excelBasedOn.WrapText;
                    newStyle.IndentLevel = excelBasedOn.IndentLevel;
                    newStyle.Orientation = excelBasedOn.Orientation;
                    newStyle.ShrinkToFit = excelBasedOn.ShrinkToFit;
                    newStyle.MergeCells = excelBasedOn.MergeCells;
                    newStyle.Locked = excelBasedOn.Locked;
                    newStyle.FormulaHidden = excelBasedOn.FormulaHidden;
                }
            }
            return newStyle;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 批量添加样式
    /// </summary>
    /// <param name="styleNames">样式名称数组</param>
    /// <returns>成功添加的样式数量</returns>
    public int AddRange(string[] styleNames)
    {
        if (_styles == null || styleNames == null || styleNames.Length == 0)
            return 0;

        int successCount = 0;
        foreach (string name in styleNames)
        {
            if (Add(name) != null)
                successCount++;
        }
        return successCount;
    }

    /// <summary>
    /// 创建内置样式
    /// </summary>
    /// <param name="builtinStyle">内置样式类型</param>
    /// <param name="name">样式名称</param>
    /// <returns>创建的样式对象</returns>
    public IExcelStyle CreateBuiltinStyle(int builtinStyle, string name = "")
    {
        if (_styles == null)
            return null;

        try
        {
            string styleName = !string.IsNullOrEmpty(name) ? name : $"BuiltinStyle{builtinStyle}";
            var style = Add(styleName);
            if (style != null)
            {
                // 设置内置样式属性（如果支持）
                // 注意：Excel中通常不能直接设置内置样式
            }
            return style;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 从模板创建工作簿样式
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <param name="includeFonts">是否包含字体</param>
    /// <param name="includeColors">是否包含颜色</param>
    /// <param name="includeBorders">是否包含边框</param>
    /// <returns>成功导入的样式数量</returns>
    public int ImportFromTemplate(string templatePath, bool includeFonts = true,
                                 bool includeColors = true, bool includeBorders = true)
    {
        if (_styles == null || string.IsNullOrEmpty(templatePath))
            return 0;

        try
        {
            // 注意：Excel Styles不直接支持从模板导入
            // 这里提供一个示例实现框架
            return 0;
        }
        catch
        {
            return 0;
        }
    }

    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据名称查找样式
    /// </summary>
    /// <param name="name">样式名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的样式数组</returns>
    public IExcelStyle[] FindByName(string name, bool matchCase = false)
    {
        if (_styles == null || string.IsNullOrEmpty(name) || Count == 0)
            return [];

        List<IExcelStyle> result = [];
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var style = this[i];
                if (style != null && style.Name != null)
                {
                    bool match = matchCase ?
                        style.Name.Contains(name) :
                        style.Name.ToLower().Contains(name.ToLower());

                    if (match)
                        result.Add(style);
                }
            }
            catch
            {
                // 忽略单个样式访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据字体查找样式
    /// </summary>
    /// <param name="fontName">字体名称</param>
    /// <param name="fontSize">字体大小</param>
    /// <param name="bold">是否粗体</param>
    /// <param name="italic">是否斜体</param>
    /// <returns>匹配的样式数组</returns>
    public IExcelStyle[] FindByFont(string fontName = "", double fontSize = 0,
                                  bool bold = false, bool italic = false)
    {
        if (_styles == null || Count == 0)
            return new IExcelStyle[0];

        var result = new System.Collections.Generic.List<IExcelStyle>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var style = this[i];
                if (style != null)
                {
                    bool match = true;

                    if (!string.IsNullOrEmpty(fontName) &&
                        style.Font?.Name?.Contains(fontName) != true)
                        match = false;

                    if (fontSize > 0 && Math.Abs(style.Font?.Size ?? 0 - fontSize) > 0.1)
                        match = false;

                    if (bold && !(style.Font?.Bold ?? false))
                        match = false;

                    if (italic && !(style.Font?.Italic ?? false))
                        match = false;

                    if (match)
                        result.Add(style);
                }
            }
            catch
            {
                // 忽略单个样式访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据颜色查找样式
    /// </summary>
    /// <param name="foregroundColor">前景色</param>
    /// <param name="backgroundColor">背景色</param>
    /// <param name="pattern">图案类型</param>
    /// <returns>匹配的样式数组</returns>
    public IExcelStyle[] FindByColor(Color? foregroundColor = null, Color? backgroundColor = null, int pattern = -1)
    {
        if (_styles == null || Count == 0)
            return [];

        var result = new System.Collections.Generic.List<IExcelStyle>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var style = this[i];
                if (style != null)
                {
                    bool match = true;

                    if (foregroundColor != null && style.Font?.Color != foregroundColor)
                        match = false;

                    if (backgroundColor != null && style.Interior?.Color != backgroundColor)
                        match = false;

                    if (pattern != -1 && style.Interior?.Pattern != pattern)
                        match = false;

                    if (match)
                        result.Add(style);
                }
            }
            catch
            {
                // 忽略单个样式访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据边框查找样式
    /// </summary>
    /// <param name="borderStyle">边框样式</param>
    /// <param name="borderColor">边框颜色</param>
    /// <param name="borderWeight">边框粗细</param>
    /// <returns>匹配的样式数组</returns>
    public IExcelStyle[] FindByBorder(int borderStyle = -1, int borderColor = -1, int borderWeight = -1)
    {
        if (_styles == null || Count == 0)
            return new IExcelStyle[0];

        var result = new System.Collections.Generic.List<IExcelStyle>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var style = this[i];
                if (style != null)
                {
                    bool match = true;

                    // 注意：边框属性访问需要更复杂的逻辑
                    // 这里提供一个简化的实现
                    if (match)
                        result.Add(style);
                }
            }
            catch
            {
                // 忽略单个样式访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取内置样式
    /// </summary>
    /// <returns>内置样式数组</returns>
    public IExcelStyle[] GetBuiltInStyles()
    {
        if (_styles == null || Count == 0)
            return new IExcelStyle[0];

        var result = new System.Collections.Generic.List<IExcelStyle>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var style = this[i];
                if (style != null && style.BuiltIn)
                {
                    result.Add(style);
                }
            }
            catch
            {
                // 忽略单个样式访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取自定义样式
    /// </summary>
    /// <returns>自定义样式数组</returns>
    public IExcelStyle[] GetCustomStyles()
    {
        if (_styles == null || Count == 0)
            return new IExcelStyle[0];

        var result = new System.Collections.Generic.List<IExcelStyle>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var style = this[i];
                if (style != null && !style.BuiltIn)
                {
                    result.Add(style);
                }
            }
            catch
            {
                // 忽略单个样式访问异常
            }
        }
        return result.ToArray();
    }
    #endregion

    #region 操作方法

    /// <summary>
    /// 删除所有自定义样式
    /// </summary>
    public void Clear()
    {
        if (_styles == null) return;

        try
        {
            // 从后往前删除，避免索引变化问题
            for (int i = Count; i >= 1; i--)
            {
                try
                {
                    var style = this[i];
                    if (style != null && !style.BuiltIn)
                    {
                        style.Delete();
                    }
                }
                catch
                {
                    // 忽略删除过程中的异常
                }
            }
        }
        catch
        {
            // 忽略清空过程中的异常
        }
    }

    /// <summary>
    /// 删除指定索引的样式
    /// </summary>
    /// <param name="index">要删除的样式索引</param>
    public void Delete(int index)
    {
        if (_styles == null || index < 1 || index > Count)
            return;

        try
        {
            _styles[index].Delete();
        }
        catch
        {
            // 忽略删除过程中的异常
        }
    }

    /// <summary>
    /// 删除指定名称的样式
    /// </summary>
    /// <param name="name">要删除的样式名称</param>
    public void Delete(string name)
    {
        if (_styles == null || string.IsNullOrEmpty(name))
            return;

        try
        {
            var style = _styles[name] as MsExcel.Style;
            style?.Delete();
        }
        catch
        {
            // 忽略删除过程中的异常
        }
    }

    /// <summary>
    /// 删除指定的样式对象
    /// </summary>
    /// <param name="style">要删除的样式对象</param>
    public void Delete(IExcelStyle style)
    {
        if (_styles == null || style == null)
            return;

        try
        {
            style.Delete();
        }
        catch
        {
            // 忽略删除过程中的异常
        }
    }

    /// <summary>
    /// 批量删除样式
    /// </summary>
    /// <param name="names">要删除的样式名称数组</param>
    public void DeleteRange(string[] names)
    {
        if (_styles == null || names == null || names.Length == 0)
            return;

        foreach (string name in names)
        {
            Delete(name);
        }
    }

    /// <summary>
    /// 重命名样式
    /// </summary>
    /// <param name="oldName">旧样式名称</param>
    /// <param name="newName">新样式名称</param>
    /// <returns>是否重命名成功</returns>
    public bool Rename(string oldName, string newName)
    {
        if (_styles == null || string.IsNullOrEmpty(oldName) || string.IsNullOrEmpty(newName))
            return false;

        try
        {
            var style = _styles[oldName] as MsExcel.Style;
            if (style != null)
            {
                // 通过复制并删除实现重命名
                var newStyle = Add(newName);
                if (newStyle != null)
                {
                    // 复制样式属性
                    var excelStyle = newStyle as ExcelStyle;
                    var oldExcelStyle = new ExcelStyle(style);

                    // 复制各种属性
                    excelStyle.Font.Name = oldExcelStyle.Font.Name;
                    excelStyle.Font.Size = oldExcelStyle.Font.Size;
                    excelStyle.Font.Bold = oldExcelStyle.Font.Bold;
                    excelStyle.Font.Italic = oldExcelStyle.Font.Italic;
                    excelStyle.Font.Color = oldExcelStyle.Font.Color;
                    excelStyle.NumberFormat = oldExcelStyle.NumberFormat;
                    excelStyle.HorizontalAlignment = oldExcelStyle.HorizontalAlignment;
                    excelStyle.VerticalAlignment = oldExcelStyle.VerticalAlignment;
                    excelStyle.WrapText = oldExcelStyle.WrapText;
                    excelStyle.IndentLevel = oldExcelStyle.IndentLevel;
                    excelStyle.Orientation = oldExcelStyle.Orientation;
                    excelStyle.ShrinkToFit = oldExcelStyle.ShrinkToFit;
                    excelStyle.MergeCells = oldExcelStyle.MergeCells;
                    excelStyle.Locked = oldExcelStyle.Locked;
                    excelStyle.FormulaHidden = oldExcelStyle.FormulaHidden;

                    // 删除原样式
                    style.Delete();
                    return true;
                }
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 复制样式
    /// </summary>
    /// <param name="sourceStyle">源样式</param>
    /// <param name="targetName">目标样式名称</param>
    /// <returns>复制的样式对象</returns>
    public IExcelStyle Copy(IExcelStyle sourceStyle, string targetName)
    {
        if (_styles == null || sourceStyle == null || string.IsNullOrEmpty(targetName))
            return null;

        try
        {
            return AddBasedOn(targetName, sourceStyle);
        }
        catch
        {
            return null;
        }
    }
    #endregion




    #region 私有辅助方法
    public IEnumerator<IExcelStyle> GetEnumerator()
    {
        for (var i = 0; i < Count; i++)
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

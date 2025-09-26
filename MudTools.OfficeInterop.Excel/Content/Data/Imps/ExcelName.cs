//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Name 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Name 对象的安全访问和资源管理
/// </summary>
internal class ExcelName : IExcelName
{
    /// <summary>
    /// 底层的 COM Name 对象
    /// </summary>
    private MsExcel.Name? _name;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelName 实例
    /// </summary>
    /// <param name="name">底层的 COM Name 对象</param>
    internal ExcelName(MsExcel.Name name)
    {
        _name = name ?? throw new ArgumentNullException(nameof(name));
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
            // 释放引用区域对象
            _refersToRange?.Dispose();

            // 释放底层COM对象
            if (_name != null)
                Marshal.ReleaseComObject(_name);
            _name = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    public string Value
    {
        get => _name?.Value;
        set
        {
            if (_name != null && value != null)
                _name.Value = value;
        }
    }

    /// <summary>
    /// 获取或设置名称
    /// </summary>
    public string Name
    {
        get => _name?.Name;
        set
        {
            if (_name != null && value != null)
                _name.Name = value;
        }
    }

    /// <summary>
    /// 获取或设置本地名称
    /// </summary>
    public string NameLocal
    {
        get => _name?.NameLocal?.ToString();
        set
        {
            if (_name != null && value != null)
                _name.NameLocal = value;
        }
    }

    /// <summary>
    /// 获取名称的索引位置
    /// </summary>
    public int Index => _name?.Index ?? 0;

    /// <summary>
    /// 获取或设置引用
    /// </summary>
    public string RefersTo
    {
        get => _name?.RefersTo?.ToString();
        set
        {
            if (_name != null && value != null)
                _name.RefersTo = value;
        }
    }

    /// <summary>
    /// 获取或设置本地引用
    /// </summary>
    public string RefersToLocal
    {
        get => _name?.RefersToLocal?.ToString();
        set
        {
            if (_name != null && value != null)
                _name.RefersToLocal = value;
        }
    }

    /// <summary>
    /// 获取或设置R1C1引用
    /// </summary>
    public string RefersToR1C1
    {
        get => _name?.RefersToR1C1?.ToString();
        set
        {
            if (_name != null && value != null)
                _name.RefersToR1C1 = value;
        }
    }

    /// <summary>
    /// 获取或设置本地R1C1引用
    /// </summary>
    public string RefersToR1C1Local
    {
        get => _name?.RefersToR1C1Local?.ToString();
        set
        {
            if (_name != null && value != null)
                _name.RefersToR1C1Local = value;
        }
    }

    /// <summary>
    /// 获取或设置是否可见
    /// </summary>
    public bool Visible
    {
        get => _name != null && _name.Visible;
        set
        {
            if (_name != null)
                _name.Visible = value;
        }
    }

    /// <summary>
    /// 获取或设置类别
    /// </summary>
    public string? Category
    {
        get => _name?.Category;
        set
        {
            if (_name != null)
                _name.Category = value;
        }
    }

    /// <summary>
    /// 获取或设置本地类别
    /// </summary>
    public string? CategoryLocal
    {
        get => _name?.CategoryLocal;
        set
        {
            if (_name != null)
                _name.CategoryLocal = value;
        }
    }

    /// <summary>
    /// 获取或设置宏类型
    /// </summary>
    public XlXLMMacroType? MacroType
    {
        get => _name?.MacroType.EnumConvert(XlXLMMacroType.xlFunction);
        set
        {
            if (_name != null)
                _name.MacroType = value.EnumConvert(MsExcel.XlXLMMacroType.xlFunction);
        }
    }

    /// <summary>
    /// 获取或设置快捷键
    /// </summary>
    public string ShortcutKey
    {
        get => _name?.ShortcutKey?.ToString();
        set
        {
            if (_name != null && value != null)
                _name.ShortcutKey = value;
        }
    }

    /// <summary>
    /// 获取名称所在的父对象
    /// </summary>
    public object Parent => _name?.Parent;

    /// <summary>
    /// 获取名称所在的Application对象
    /// </summary>
    public IExcelApplication Application
    {
        get
        {
            var application = _name?.Application as MsExcel.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    /// <summary>
    /// 引用区域对象缓存
    /// </summary>
    private IExcelRange _refersToRange;

    /// <summary>
    /// 获取引用的区域对象
    /// </summary>
    public IExcelRange RefersToRange => _refersToRange ??= new ExcelRange(_name?.RefersToRange as MsExcel.Range);

    /// <summary>
    /// 获取或设置注释
    /// </summary>
    public string Comment
    {
        get => _name?.Comment?.ToString();
        set
        {
            if (_name != null && value != null)
                _name.Comment = value;
        }
    }

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除名称
    /// </summary>
    public void Delete()
    {
        _name?.Delete();
    }

    /// <summary>
    /// 选择名称引用的区域
    /// </summary>
    public void Select()
    {
        try
        {
            _name?.RefersToRange?.Select();
        }
        catch
        {
            // 忽略选择过程中的异常
        }
    }

    /// <summary>
    /// 激活名称引用的区域
    /// </summary>
    public void Activate()
    {
        try
        {
            _name?.RefersToRange?.Activate();
        }
        catch
        {
            // 忽略激活过程中的异常
        }
    }

    /// <summary>
    /// 复制名称
    /// </summary>
    /// <param name="newName">新名称</param>
    /// <param name="parent">父对象</param>
    /// <returns>复制的名称对象</returns>
    public IExcelName? Copy(string newName = "", object? parent = null)
    {
        if (_name?.Parent == null)
            return null;

        try
        {
            string copyName = !string.IsNullOrEmpty(newName) ? newName : $"{_name.Name}_Copy";
            var namesCollection = _name.Parent as MsExcel.Names;

            if (namesCollection != null)
            {
                var copiedName = namesCollection.Add(
                    copyName,
                    _name.RefersTo,
                    _name.Visible,
                    _name.MacroType,
                    _name.ShortcutKey,
                    _name.Category,
                    _name.NameLocal,
                    _name.RefersToLocal,
                    _name.CategoryLocal,
                    _name.RefersToR1C1,
                    _name.RefersToR1C1Local
                ) as MsExcel.Name;

                return copiedName != null ? new ExcelName(copiedName) : null;
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 重命名名称
    /// </summary>
    /// <param name="newName">新名称</param>
    public void Rename(string newName)
    {
        if (_name == null || string.IsNullOrEmpty(newName))
            return;

        try
        {
            // 通过复制并删除实现重命名
            var parent = _name.Parent as MsExcel.Names;
            if (parent != null)
            {
                var newNameObject = parent.Add(
                    newName,
                    _name.RefersTo,
                    _name.Visible,
                    _name.MacroType,
                    _name.ShortcutKey,
                    _name.Category,
                    _name.NameLocal,
                    _name.RefersToLocal,
                    _name.CategoryLocal,
                    _name.RefersToR1C1,
                    _name.RefersToR1C1Local
                ) as MsExcel.Name;

                if (newNameObject != null)
                {
                    _name.Delete();
                    _name = newNameObject;
                }
            }
        }
        catch
        {
            // 忽略重命名过程中的异常
        }
    }

    /// <summary>
    /// 更新引用
    /// </summary>
    /// <param name="newRefersTo">新引用</param>
    public void UpdateRefersTo(string newRefersTo)
    {
        if (_name == null || string.IsNullOrEmpty(newRefersTo))
            return;

        try
        {
            _name.RefersTo = newRefersTo;
        }
        catch
        {
            // 忽略更新过程中的异常
        }
    }

    /// <summary>
    /// 刷新名称
    /// </summary>
    public void Refresh()
    {
        // Excel名称通常会自动刷新
        // 这里提供一个空实现以保持接口一致性
    }

    #endregion

    #region 高级功能    

    /// <summary>
    /// 转换引用格式
    /// </summary>
    /// <param name="toR1C1">是否转换为R1C1格式</param>
    /// <returns>转换后的引用</returns>
    public string ConvertReferenceFormat(bool toR1C1 = true)
    {
        if (_name == null)
            return "";

        try
        {
            if (toR1C1)
                return (string)(_name.RefersToR1C1 ?? "");
            else
                return (string)(_name.RefersTo ?? "");
        }
        catch
        {
            return "";
        }
    }


    /// <summary>
    /// 检查循环引用
    /// </summary>
    /// <returns>是否存在循环引用</returns>
    public bool HasCircularReference()
    {
        // 注意：Excel COM对象不直接提供循环引用检查
        // 这里提供一个基础实现
        try
        {
            // 简单检查引用是否包含自身名称
            if (!string.IsNullOrEmpty(RefersTo) && !string.IsNullOrEmpty(Name))
            {
                return RefersTo.Contains(Name);
            }
            return false;
        }
        catch
        {
            return false;
        }
    }
    #endregion
}
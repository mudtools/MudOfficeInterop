//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// ControlFormat COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class ExcelControlFormat : IExcelControlFormat
{
    /// <summary>
    /// 内部持有的原始 COM 对象。
    /// </summary>
    internal MsExcel.ControlFormat _controlFormat;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="controlFormat">原始的 ControlFormat COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 controlFormat 为 null 时抛出。</exception>
    internal ExcelControlFormat(MsExcel.ControlFormat controlFormat)
    {
        _controlFormat = controlFormat ?? throw new ArgumentNullException(nameof(controlFormat));
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的受保护虚方法，支持派生类重写。
    /// </summary>
    /// <param name="disposing">是否由用户代码显式调用释放。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放托管资源：释放 COM 对象
            if (_controlFormat != null)
            {
                Marshal.ReleaseComObject(_controlFormat);
                _controlFormat = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 公开的 Dispose 方法，用于显式释放资源。
    /// 调用后对象不应再被使用。
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取此对象的父对象（通常是 Shape）。
    /// </summary>
    public object Parent => _controlFormat?.Parent;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication Application =>
        _controlFormat?.Application != null
            ? new ExcelApplication(_controlFormat.Application as MsExcel.Application)
            : null;

    /// <summary>
    /// 获取或设置控件的当前选中项索引（从 1 开始）。
    /// </summary>
    public int Value
    {
        get => _controlFormat?.Value ?? 0;
        set
        {
            if (_controlFormat != null)
                _controlFormat.Value = value;
        }
    }

    /// <summary>
    /// 获取或设置控件允许的最小值（适用于滚动条、微调项等）。
    /// </summary>
    public int Min
    {
        get => _controlFormat?.Min ?? 0;
        set
        {
            if (_controlFormat != null)
                _controlFormat.Min = value;
        }
    }

    /// <summary>
    /// 获取或设置控件允许的最大值（适用于滚动条、微调项等）。
    /// </summary>
    public int Max
    {
        get => _controlFormat?.Max ?? 0;
        set
        {
            if (_controlFormat != null)
                _controlFormat.Max = value;
        }
    }


    /// <summary>
    /// 获取或设置控件是否允许多选（适用于列表框）。
    /// </summary>
    public bool MultiSelect
    {
        get => _controlFormat != null && _controlFormat.MultiSelect.ConvertToBool();
        set
        {
            if (_controlFormat != null)
                _controlFormat.MultiSelect = value ? 1 : 0;
        }
    }

    /// <summary>
    /// 获取控件中列表项的总数。
    /// </summary>
    public int ListCount => _controlFormat?.ListCount ?? 0;

    public int ListIndex
    {
        get => _controlFormat?.ListIndex ?? 0;
        set
        {
            if (_controlFormat != null)
                _controlFormat.ListIndex = value;
        }
    }

    public int SmallChange
    {
        get => _controlFormat?.SmallChange ?? 0;
        set
        {
            if (_controlFormat != null)
                _controlFormat.SmallChange = value;
        }

    }

    public bool LockedText
    {
        get => _controlFormat != null && _controlFormat.LockedText;
        set
        {
            if (_controlFormat != null)
                _controlFormat.LockedText = value;
        }
    }

    public bool PrintObject
    {
        get => _controlFormat != null && _controlFormat.PrintObject;
        set
        {
            if (_controlFormat != null)
                _controlFormat.PrintObject = value;
        }
    }

    /// <summary>
    /// 获取或设置与控件关联的数据源区域（用于动态填充列表项）。
    /// </summary>
    public string ListFillRange
    {
        get => _controlFormat?.ListFillRange != null
            ? _controlFormat.ListFillRange
            : string.Empty;

        set
        {
            if (_controlFormat != null && value != null)
            {
                _controlFormat.ListFillRange = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置与控件值绑定的单元格（控件值变化时自动写入该单元格）。
    /// </summary>
    public string LinkedCell
    {
        get => _controlFormat?.LinkedCell != null
            ? _controlFormat.LinkedCell
            : string.Empty;

        set
        {
            if (_controlFormat != null && value != null)
            {
                _controlFormat.LinkedCell = value;
            }
        }
    }

    /// <summary>
    /// 向控件列表中添加一个新项。
    /// </summary>
    /// <param name="text">要添加的文本。</param>
    /// <param name="index">插入位置（从1开始），0=追加到末尾。</param>
    public void AddItem(string text, int index = 0)
    {
        if (_controlFormat == null || string.IsNullOrEmpty(text)) return;
        _controlFormat.AddItem(text, index);
    }

    /// <summary>
    /// 从控件列表中移除指定索引的项。
    /// </summary>
    /// <param name="index">要删除的项索引（从1开始）。</param>
    public void RemoveItem(int index)
    {
        if (_controlFormat == null || index < 1) return;
        _controlFormat.RemoveItem(index);
    }

    /// <summary>
    /// 清空控件中的所有列表项。
    /// </summary>
    public void RemoveAllItems()
    {
        _controlFormat?.RemoveAllItems();
    }

    /// <summary>
    /// 获取指定索引项的文本内容。
    /// </summary>
    /// <param name="index">项索引（从1开始）。</param>
    /// <returns>项文本，若无效则返回空字符串。</returns>
    public string GetItemText(int index)
    {
        if (_controlFormat == null || index < 1) return string.Empty;
        try
        {
            return _controlFormat.List[index] as string ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }
}
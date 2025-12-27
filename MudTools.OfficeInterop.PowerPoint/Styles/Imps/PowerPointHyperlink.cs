//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 超链接实现类
/// </summary>
internal class PowerPointHyperlink : IPowerPointHyperlink
{
    private readonly MsPowerPoint.Hyperlink _hyperlink;
    private bool _disposedValue;

    /// <summary>
    /// 获取或设置超链接地址
    /// </summary>
    public string Address
    {
        get => _hyperlink?.Address ?? string.Empty;
        set
        {
            if (_hyperlink != null)
                _hyperlink.Address = value ?? string.Empty;
        }
    }

    /// <summary>
    /// 获取或设置子地址
    /// </summary>
    public string SubAddress
    {
        get => _hyperlink?.SubAddress ?? string.Empty;
        set
        {
            if (_hyperlink != null)
                _hyperlink.SubAddress = value ?? string.Empty;
        }
    }

    /// <summary>
    /// 获取或设置显示文本
    /// </summary>
    public string TextToDisplay
    {
        get => _hyperlink?.TextToDisplay ?? string.Empty;
        set
        {
            if (_hyperlink != null)
                _hyperlink.TextToDisplay = value ?? string.Empty;
        }
    }

    /// <summary>
    /// 获取或设置屏幕提示
    /// </summary>
    public string ScreenTip
    {
        get => _hyperlink?.ScreenTip ?? string.Empty;
        set
        {
            if (_hyperlink != null)
                _hyperlink.ScreenTip = value ?? string.Empty;
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _hyperlink?.Parent;

    /// <summary>
    /// 获取超链接类型
    /// </summary>
    public int Type => _hyperlink != null ? (int)_hyperlink.Type : 0;

    /// <summary>
    /// 获取超链接是否有效
    /// </summary>
    public bool IsValid
    {
        get
        {
            try
            {
                return !string.IsNullOrEmpty(Address) || !string.IsNullOrEmpty(SubAddress);
            }
            catch
            {
                return false;
            }
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="hyperlink">COM Hyperlink 对象</param>
    internal PowerPointHyperlink(MsPowerPoint.Hyperlink hyperlink)
    {
        _hyperlink = hyperlink ?? throw new ArgumentNullException(nameof(hyperlink));
        _disposedValue = false;
    }

    /// <summary>
    /// 跟随超链接
    /// </summary>
    public void Follow()
    {
        try
        {
            _hyperlink?.Follow();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to follow hyperlink.", ex);
        }
    }

    /// <summary>
    /// 删除超链接
    /// </summary>
    public void Delete()
    {
        try
        {
            _hyperlink?.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete hyperlink.", ex);
        }
    }

    /// <summary>
    /// 编辑超链接
    /// </summary>
    /// <param name="newAddress">新地址</param>
    /// <param name="newSubAddress">新子地址</param>
    /// <param name="newTextToDisplay">新显示文本</param>
    public void Edit(string newAddress = null, string newSubAddress = null, string newTextToDisplay = null)
    {
        try
        {
            if (_hyperlink != null)
            {
                if (!string.IsNullOrEmpty(newAddress))
                    _hyperlink.Address = newAddress;
                if (!string.IsNullOrEmpty(newSubAddress))
                    _hyperlink.SubAddress = newSubAddress;
                if (!string.IsNullOrEmpty(newTextToDisplay))
                    _hyperlink.TextToDisplay = newTextToDisplay;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to edit hyperlink.", ex);
        }
    }

    /// <summary>
    /// 复制超链接
    /// </summary>
    /// <returns>复制的超链接对象</returns>
    public IPowerPointHyperlink Duplicate()
    {
        try
        {
            // PowerPoint 中没有直接的复制超链接方法
            throw new NotImplementedException("Duplicating hyperlink is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to duplicate hyperlink.", ex);
        }
    }

    /// <summary>
    /// 应用超链接到指定范围
    /// </summary>
    /// <param name="range">目标范围</param>
    public void ApplyTo(object range)
    {
        if (range == null)
            throw new ArgumentNullException(nameof(range));

        try
        {
            // 这需要具体的实现来应用超链接到范围
            throw new NotImplementedException("Applying hyperlink to range is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply hyperlink to range.", ex);
        }
    }

    /// <summary>
    /// 验证超链接
    /// </summary>
    /// <returns>是否有效</returns>
    public bool Validate()
    {
        try
        {
            return IsValid;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to validate hyperlink.", ex);
        }
    }

    /// <summary>
    /// 获取超链接信息
    /// </summary>
    /// <returns>超链接信息字符串</returns>
    public string GetHyperlinkInfo()
    {
        try
        {
            return $"Hyperlink: {TextToDisplay} -> {Address}{(!string.IsNullOrEmpty(SubAddress) ? "#" + SubAddress : "")}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get hyperlink info.", ex);
        }
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Interop.PowerPoint;

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint 颜色方案实现类
/// </summary>
internal class PowerPointColorScheme : IPowerPointColorScheme
{
    private readonly MsPowerPoint.ColorScheme _colorScheme;
    private bool _disposedValue;

    /// <summary>
    /// 获取颜色数量
    /// </summary>
    public int Count => _colorScheme?.Count ?? 0;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _colorScheme?.Parent;

    /// <summary>
    /// 获取或设置指定索引的颜色
    /// </summary>
    /// <param name="schemeColorIndex">颜色方案索引</param>
    /// <returns>颜色值</returns>
    public int this[int schemeColorIndex]
    {
        get => Colors(schemeColorIndex);
        set => SetColor(schemeColorIndex, value);
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="colorScheme">COM ColorScheme 对象</param>
    internal PowerPointColorScheme(MsPowerPoint.ColorScheme colorScheme)
    {
        _colorScheme = colorScheme;
        _disposedValue = false;
    }

    /// <summary>
    /// 获取指定索引的颜色
    /// </summary>
    /// <param name="index">颜色索引</param>
    /// <returns>颜色值</returns>
    public int Colors(int index)
    {
        try
        {
            if (_colorScheme != null && index >= 1 && index <= Count)
            {
                return _colorScheme[(PpColorSchemeIndex)index].RGB;
            }
            return 0;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get color at index {index}.", ex);
        }
    }

    /// <summary>
    /// 获取指定名称的颜色
    /// </summary>
    /// <param name="name">颜色名称</param>
    /// <returns>颜色值</returns>
    public int Colors(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Color name cannot be null or empty.", nameof(name));

        try
        {
            // PowerPoint ColorScheme 不支持按名称获取颜色
            throw new NotImplementedException("Getting color by name is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get color by name '{name}'.", ex);
        }
    }

    /// <summary>
    /// 应用颜色方案
    /// </summary>
    /// <param name="schemeIndex">方案索引</param>
    public void Apply(int schemeIndex)
    {
        try
        {
            // 颜色方案应用需要通过演示文稿对象实现
            throw new NotImplementedException("Applying color scheme is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply color scheme.", ex);
        }
    }

    /// <summary>
    /// 保存颜色方案
    /// </summary>
    /// <param name="fileName">文件路径</param>
    public void Save(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

        try
        {
            // PowerPoint 不直接支持保存颜色方案
            throw new NotImplementedException("Saving color scheme is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to save color scheme.", ex);
        }
    }

    /// <summary>
    /// 加载颜色方案
    /// </summary>
    /// <param name="fileName">文件路径</param>
    public void Load(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

        if (!System.IO.File.Exists(fileName))
            throw new System.IO.FileNotFoundException("Color scheme file not found.", fileName);

        try
        {
            // PowerPoint 不直接支持加载颜色方案
            throw new NotImplementedException("Loading color scheme is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to load color scheme.", ex);
        }
    }

    /// <summary>
    /// 重置颜色方案
    /// </summary>
    public void Reset()
    {
        try
        {
            // 颜色方案重置需要通过演示文稿对象实现
            throw new NotImplementedException("Resetting color scheme is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset color scheme.", ex);
        }
    }


    /// <summary>
    /// 设置颜色值
    /// </summary>
    /// <param name="index">颜色索引</param>
    /// <param name="color">颜色值</param>
    public void SetColor(int index, int color)
    {
        try
        {
            if (_colorScheme != null && index >= 1 && index <= Count)
            {
                _colorScheme[(PpColorSchemeIndex)index].RGB = color;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to set color at index {index}.", ex);
        }
    }

    /// <summary>
    /// 应用到所有幻灯片
    /// </summary>
    public void ApplyToAll()
    {
        try
        {
            // 颜色方案应用到所有幻灯片需要通过演示文稿对象实现
            throw new NotImplementedException("Applying color scheme to all slides is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply color scheme to all slides.", ex);
        }
    }

    /// <summary>
    /// 获取颜色方案信息
    /// </summary>
    /// <returns>颜色方案信息字符串</returns>
    public string GetColorSchemeInfo()
    {
        try
        {
            return $"ColorScheme - Count: {Count}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get color scheme info.", ex);
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

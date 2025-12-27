//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 声音效果实现类
/// </summary>
internal class PowerPointSoundEffect : IPowerPointSoundEffect
{
    private readonly MsPowerPoint.SoundEffect _soundEffect;
    private bool _disposedValue;

    /// <summary>
    /// 获取或设置声音名称
    /// </summary>
    public string Name
    {
        get => _soundEffect?.Name ?? string.Empty;
        set
        {
            if (_soundEffect != null)
                _soundEffect.Name = value ?? string.Empty;
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _soundEffect?.Parent;


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="soundEffect">COM SoundEffect 对象</param>
    internal PowerPointSoundEffect(MsPowerPoint.SoundEffect soundEffect)
    {
        _soundEffect = soundEffect;
        _disposedValue = false;
    }

    /// <summary>
    /// 从文件导入声音
    /// </summary>
    /// <param name="fileName">文件路径</param>
    public void ImportFromFile(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

        if (!System.IO.File.Exists(fileName))
            throw new System.IO.FileNotFoundException("Sound file not found.", fileName);

        try
        {
            _soundEffect?.ImportFromFile(fileName);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to import sound from '{fileName}'.", ex);
        }
    }

    /// <summary>
    /// 播放声音
    /// </summary>
    public void Play()
    {
        try
        {
            _soundEffect?.Play();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to play sound.", ex);
        }
    }

    /// <summary>
    /// 停止播放
    /// </summary>
    public void Stop()
    {
        try
        {
            // PowerPoint SoundEffect 没有直接的停止方法
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to stop sound.", ex);
        }
    }

    /// <summary>
    /// 暂停播放
    /// </summary>
    public void Pause()
    {
        try
        {
            // PowerPoint SoundEffect 没有直接的暂停方法
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to pause sound.", ex);
        }
    }

    /// <summary>
    /// 恢复播放
    /// </summary>
    public void Resume()
    {
        try
        {
            // PowerPoint SoundEffect 没有直接的恢复方法
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to resume sound.", ex);
        }
    }

    /// <summary>
    /// 删除声音效果
    /// </summary>
    public void Delete()
    {
        try
        {
            // PowerPoint SoundEffect 没有直接的删除方法
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete sound effect.", ex);
        }
    }

    /// <summary>
    /// 获取声音效果信息
    /// </summary>
    /// <returns>声音效果信息字符串</returns>
    public string GetSoundEffectInfo()
    {
        try
        {
            return $"SoundEffect - Name: {Name}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get sound effect info.", ex);
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

//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;


/// <summary>
/// PowerPoint 动作设置集合实现类
/// </summary>
internal class PowerPointActionSettings : IPowerPointActionSettings
{
    private readonly MsPowerPoint.ActionSettings _actionSettings;
    private bool _disposedValue;

    /// <summary>
    /// 获取动作设置数量
    /// </summary>
    public int Count => _actionSettings?.Count ?? 0;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _actionSettings?.Parent;

    /// <summary>
    /// 根据索引获取动作设置
    /// </summary>
    public IPowerPointActionSetting this[int index]
    {
        get
        {
            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

            try
            {
                var actionSetting = _actionSettings[(MsPowerPoint.PpMouseActivation)index];
                return actionSetting != null ? new PowerPointActionSetting(actionSetting) : null;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get action setting at index {index}.", ex);
            }
        }
    }

    /// <summary>
    /// 根据动作类型获取动作设置
    /// </summary>
    public IPowerPointActionSetting this[PpMouseActivation actionType]
    {
        get
        {
            try
            {
                var actionSetting = _actionSettings[(MsPowerPoint.PpMouseActivation)actionType];
                return actionSetting != null ? new PowerPointActionSetting(actionSetting) : null;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get action setting for action type {actionType}.", ex);
            }
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="actionSettings">COM ActionSettings 对象</param>
    internal PowerPointActionSettings(MsPowerPoint.ActionSettings actionSettings)
    {
        _actionSettings = actionSettings;
        _disposedValue = false;
    }





    /// <summary>
    /// 查找符合条件的动作设置
    /// </summary>
    /// <param name="predicate">查找条件</param>
    /// <returns>符合条件的动作设置列表</returns>
    public IEnumerable<IPowerPointActionSetting> Find(Func<IPowerPointActionSetting, bool> predicate)
    {
        if (predicate == null)
            throw new ArgumentNullException(nameof(predicate));

        var results = new List<IPowerPointActionSetting>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var actionSetting = this[i];
                    if (actionSetting != null && predicate(actionSetting))
                    {
                        results.Add(actionSetting);
                    }
                }
                catch
                {
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to find action settings.", ex);
        }
        return results;
    }

    /// <summary>
    /// 获取所有鼠标点击动作设置
    /// </summary>
    /// <returns>鼠标点击动作设置列表</returns>
    public IEnumerable<IPowerPointActionSetting> GetMouseClickActions()
    {
        try
        {
            var mouseClickAction = _actionSettings[MsPowerPoint.PpMouseActivation.ppMouseClick];
            if (mouseClickAction != null)
            {
                return new List<IPowerPointActionSetting> { new PowerPointActionSetting(mouseClickAction) };
            }
            return new List<IPowerPointActionSetting>();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get mouse click actions.", ex);
        }
    }

    /// <summary>
    /// 获取所有鼠标悬停动作设置
    /// </summary>
    /// <returns>鼠标悬停动作设置列表</returns>
    public IEnumerable<IPowerPointActionSetting> GetMouseOverActions()
    {
        try
        {
            var mouseOverAction = _actionSettings[MsPowerPoint.PpMouseActivation.ppMouseOver];
            if (mouseOverAction != null)
            {
                return new List<IPowerPointActionSetting> { new PowerPointActionSetting(mouseOverAction) };
            }
            return new List<IPowerPointActionSetting>();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get mouse over actions.", ex);
        }
    }

    /// <summary>
    /// 清除所有动作设置
    /// </summary>
    public void Clear()
    {
        try
        {
            // PowerPoint ActionSettings 不能直接清除，需要重置每个动作设置
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var actionSetting = this[i];
                    actionSetting?.Reset();
                }
                catch
                {
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to clear action settings.", ex);
        }
    }

    /// <summary>
    /// 重置动作设置
    /// </summary>
    public void Reset()
    {
        try
        {
            Clear();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset action settings.", ex);
        }
    }

    /// <summary>
    /// 刷新动作设置显示
    /// </summary>
    public void Refresh()
    {
        try
        {
            // 动作设置刷新通常自动进行
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh action settings.", ex);
        }
    }

    /// <summary>
    /// 获取动作设置集合信息
    /// </summary>
    /// <returns>动作设置集合信息字符串</returns>
    public string GetActionSettingsInfo()
    {
        try
        {
            return $"ActionSettings - Count: {Count}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get action settings info.", ex);
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

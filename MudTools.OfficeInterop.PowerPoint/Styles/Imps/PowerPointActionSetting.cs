//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;



/// <summary>
/// PowerPoint 动作设置实现类
/// </summary>
internal class PowerPointActionSetting : IPowerPointActionSetting
{
    private readonly MsPowerPoint.ActionSetting _actionSetting;
    private bool _disposedValue;

    /// <summary>
    /// 获取或设置动作类型
    /// </summary>
    public PpActionType ActionType
    {
        get => _actionSetting != null ? (PpActionType)_actionSetting.Action : PpActionType.ppActionNone;
        set
        {
            if (_actionSetting != null)
                _actionSetting.Action = (MsPowerPoint.PpActionType)value;
        }
    }

    /// <summary>
    /// 获取或设置超链接
    /// </summary>
    public string Hyperlink
    {
        get => _actionSetting.Hyperlink.Address;
    }

    /// <summary>
    /// 获取或设置运行程序
    /// </summary>
    public string Run
    {
        get => _actionSetting?.Run ?? string.Empty;
        set
        {
            if (_actionSetting != null)
                _actionSetting.Run = value ?? string.Empty;
        }
    }

    /// <summary>
    /// 获取或设置幻灯片放映名称
    /// </summary>
    public string SlideShowName
    {
        get => _actionSetting?.SlideShowName ?? string.Empty;
        set
        {
            if (_actionSetting != null)
                _actionSetting.SlideShowName = value ?? string.Empty;
        }
    }

    /// <summary>
    /// 获取或设置动画动作
    /// </summary>
    public PpAnimateAction AnimateAction
    {
        get => _actionSetting != null ? (PpAnimateAction)_actionSetting.AnimateAction : PpAnimateAction.ppAnimateNone;
        set
        {
            if (_actionSetting != null)
                _actionSetting.AnimateAction = (MsCore.MsoTriState)value;
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _actionSetting?.Parent;


    /// <summary>
    /// 获取或设置触发器类型
    /// </summary>
    public PpMouseActivation TriggerType
    {
        get => _actionSetting != null ? (PpMouseActivation)_actionSetting.Parent : PpMouseActivation.ppMouseClick;
        set
        {
            // TriggerType 是只读属性，由父对象决定
        }
    }


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="actionSetting">COM ActionSetting 对象</param>
    internal PowerPointActionSetting(MsPowerPoint.ActionSetting actionSetting)
    {
        _actionSetting = actionSetting;
        _disposedValue = false;
    }

    /// <summary>
    /// 设置动作参数
    /// </summary>
    /// <param name="actionType">动作类型</param>
    /// <param name="hyperlink">超链接</param>
    /// <param name="run">运行程序</param>
    /// <param name="slideShowName">幻灯片放映名称</param>
    public void SetAction(PpActionType actionType = PpActionType.ppActionNone, string hyperlink = null, string run = null, string slideShowName = null)
    {
        try
        {
            if (_actionSetting != null)
            {
                _actionSetting.Action = (MsPowerPoint.PpActionType)actionType;
                if (!string.IsNullOrEmpty(hyperlink))
                    _actionSetting.Hyperlink.Address = hyperlink;
                if (!string.IsNullOrEmpty(run))
                    _actionSetting.Run = run;
                if (!string.IsNullOrEmpty(slideShowName))
                    _actionSetting.SlideShowName = slideShowName;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set action parameters.", ex);
        }
    }

    /// <summary>
    /// 设置动画效果
    /// </summary>
    /// <param name="animateAction">动画动作</param>
    /// <param name="playAnimation">是否播放动画</param>
    /// <param name="stopAnimation">是否停止动画</param>
    public void SetAnimation(PpAnimateAction animateAction = PpAnimateAction.ppAnimateNone, bool playAnimation = false, bool stopAnimation = false)
    {
        try
        {
            if (_actionSetting != null)
            {
                _actionSetting.AnimateAction = (MsCore.MsoTriState)animateAction;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set animation effect.", ex);
        }
    }

    /// <summary>
    /// 应用动作设置到对象
    /// </summary>
    /// <param name="targetObject">目标对象</param>
    public void ApplyTo(object targetObject)
    {
        if (targetObject == null)
            throw new ArgumentNullException(nameof(targetObject));

        try
        {
            // 动作设置通常是直接关联到对象的，不需要额外应用
            throw new NotImplementedException("Applying action setting to object is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply action setting to object.", ex);
        }
    }


    /// <summary>
    /// 预览动作
    /// </summary>
    public void Preview()
    {
        try
        {
            // 动作预览需要通过幻灯片放映实现
            throw new NotImplementedException("Previewing action is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to preview action.", ex);
        }
    }

    /// <summary>
    /// 复制动作设置
    /// </summary>
    /// <returns>复制的动作设置</returns>
    public IPowerPointActionSetting Duplicate()
    {
        try
        {
            // PowerPoint ActionSetting 没有直接的复制方法
            throw new NotImplementedException("Duplicating action setting is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to duplicate action setting.", ex);
        }
    }

    /// <summary>
    /// 重置动作设置
    /// </summary>
    public void Reset()
    {
        try
        {
            if (_actionSetting != null)
            {
                _actionSetting.Action = MsPowerPoint.PpActionType.ppActionNone;
                _actionSetting.Hyperlink.Address = string.Empty;
                _actionSetting.Run = string.Empty;
                _actionSetting.SlideShowName = string.Empty;
                _actionSetting.AnimateAction = MsCore.MsoTriState.msoFalse;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset action setting.", ex);
        }
    }

    /// <summary>
    /// 获取动作设置信息
    /// </summary>
    /// <returns>动作设置信息字符串</returns>
    public string GetActionSettingInfo()
    {
        try
        {
            return $"ActionSetting - Action: {ActionType}, Hyperlink: {Hyperlink}, Run: {Run}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get action setting info.", ex);
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

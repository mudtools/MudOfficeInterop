//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 动作设置接口
/// </summary>
public interface IPowerPointActionSetting : IDisposable
{
    /// <summary>
    /// 获取或设置动作类型
    /// </summary>
    PpActionType ActionType { get; set; }

    /// <summary>
    /// 获取或设置超链接
    /// </summary>
    string Hyperlink { get; }

    /// <summary>
    /// 获取或设置运行程序
    /// </summary>
    string Run { get; set; }

    /// <summary>
    /// 获取或设置幻灯片放映名称
    /// </summary>
    string SlideShowName { get; set; }

    /// <summary>
    /// 获取或设置动画动作
    /// </summary>
    PpAnimateAction AnimateAction { get; set; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }


    /// <summary>
    /// 获取或设置触发器类型
    /// </summary>
    PpMouseActivation TriggerType { get; set; }

    /// <summary>
    /// 设置动作参数
    /// </summary>
    /// <param name="actionType">动作类型</param>
    /// <param name="hyperlink">超链接</param>
    /// <param name="run">运行程序</param>
    /// <param name="slideShowName">幻灯片放映名称</param>
    void SetAction(PpActionType actionType = PpActionType.ppActionNone, string hyperlink = null, string run = null, string slideShowName = null);


    /// <summary>
    /// 设置动画效果
    /// </summary>
    /// <param name="animateAction">动画动作</param>
    /// <param name="playAnimation">是否播放动画</param>
    /// <param name="stopAnimation">是否停止动画</param>
    void SetAnimation(PpAnimateAction animateAction = PpAnimateAction.ppAnimateNone, bool playAnimation = false, bool stopAnimation = false);

    /// <summary>
    /// 应用动作设置到对象
    /// </summary>
    /// <param name="targetObject">目标对象</param>
    void ApplyTo(object targetObject);


    /// <summary>
    /// 预览动作
    /// </summary>
    void Preview();

    /// <summary>
    /// 复制动作设置
    /// </summary>
    /// <returns>复制的动作设置</returns>
    IPowerPointActionSetting Duplicate();

    /// <summary>
    /// 重置动作设置
    /// </summary>
    void Reset();

    /// <summary>
    /// 获取动作设置信息
    /// </summary>
    /// <returns>动作设置信息字符串</returns>
    string GetActionSettingInfo();
}
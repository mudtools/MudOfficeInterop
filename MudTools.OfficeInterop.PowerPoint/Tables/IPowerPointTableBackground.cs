//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 表格的背景格式设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointTableBackground : IDisposable
{
    /// <summary>
    /// 获取表格背景的填充格式设置。
    /// </summary>
    /// <value>表示表格背景填充格式的 <see cref="IPowerPointFillFormat"/> 对象。</value>
    IPowerPointFillFormat? Fill { get; }

    /// <summary>
    /// 获取表格背景的图片格式设置。
    /// </summary>
    /// <value>表示表格背景图片格式的 <see cref="IPowerPointPictureFormat"/> 对象。</value>
    IPowerPointPictureFormat? Picture { get; }

    /// <summary>
    /// 获取表格背景的反射效果格式设置。
    /// </summary>
    /// <value>表示表格背景反射效果的 <see cref="IOfficeReflectionFormat"/> 对象。</value>
    IOfficeReflectionFormat? Reflection { get; }

    /// <summary>
    /// 获取表格背景的阴影效果格式设置。
    /// </summary>
    /// <value>表示表格背景阴影效果的 <see cref="IPowerPointShadowFormat"/> 对象。</value>
    IPowerPointShadowFormat? Shadow { get; }
}
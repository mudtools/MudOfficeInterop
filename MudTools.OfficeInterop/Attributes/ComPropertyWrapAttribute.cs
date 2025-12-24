//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;


/// <summary>
/// COM封装接口的属性信息。
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public class ComPropertyWrapAttribute : Attribute
{
    /// <summary>
    /// 属性对象所在的COM对象命名空间。
    /// </summary>
    public string? ComNamespace { get; set; }

    /// <summary>
    /// 属性默认值。
    /// </summary>
    public string? DefaultValue { get; set; }

    /// <summary>
    /// 是否需要转换属性值。
    /// </summary>
    public bool NeedConvert { get; set; } = false;

    /// <summary>
    /// 是否需要释放属性值资源。
    /// </summary>
    public bool NeedDispose { get; set; } = true;

    /// <summary>
    /// 标记该属性采用get、set方法进行访问。
    /// </summary>
    public bool IsMethod { get; set; }

    /// <summary>
    /// 获取或设置属性名，默认为空（即原始属性名。）
    /// </summary>
    public string? PropertyName { get; set; }
}

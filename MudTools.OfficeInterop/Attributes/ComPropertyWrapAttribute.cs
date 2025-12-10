namespace MudTools.OfficeInterop;


/// <summary>
/// COM封装接口的属性信息。
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public class ComPropertyWrapAttribute : Attribute
{
    /// <summary>
    /// 属性对象类型
    /// </summary>
    public PropertyType PropertyType { get; set; } = PropertyType.ValueType;

    /// <summary>
    /// 属性默认值。
    /// </summary>
    public string? DefaultValue { get; set; }
}

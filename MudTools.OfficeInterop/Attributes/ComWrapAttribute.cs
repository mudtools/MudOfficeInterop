namespace MudTools.OfficeInterop;

/// <summary>
/// 指示封装的是一个普通的COM对象。
/// </summary>
public abstract class ComWrapAttribute : Attribute
{
    /// <summary>
    /// COM对象所在的命名空间。
    /// </summary>
    public string? ComNamespace { get; set; }

    /// <summary>
    /// COM对象类名
    /// </summary>
    public string? ComClassName { get; set; }
}

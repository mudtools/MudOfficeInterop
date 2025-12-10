namespace MudTools.OfficeInterop;

/// <summary>
/// 指示封装的是一个普通的COM对象。
/// </summary>
[AttributeUsage(AttributeTargets.Interface, AllowMultiple = false, Inherited = false)]
public class ComObjectWrapAttribute : ComWrapAttribute
{
}

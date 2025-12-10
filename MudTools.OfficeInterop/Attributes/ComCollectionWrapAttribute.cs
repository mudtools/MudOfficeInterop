namespace MudTools.OfficeInterop;

/// <summary>
/// 指示封装的是一个集合类型的COM对象。
/// </summary>
[AttributeUsage(AttributeTargets.Interface, AllowMultiple = false, Inherited = false)]
public class ComCollectionWrapAttribute : ComWrapAttribute
{
}

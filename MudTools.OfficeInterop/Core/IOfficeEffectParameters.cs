namespace MudTools.OfficeInterop;

/// <summary>
/// 封装 Microsoft.Office.Core.EffectParameters 的接口，用于操作效果参数集合。
/// </summary>
public interface IOfficeEffectParameters : IEnumerable<IOfficeEffectParameter>, IDisposable
{
    /// <summary>
    /// 获取效果参数集合中的参数数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取效果参数（从1开始）。
    /// </summary>
    IOfficeEffectParameter this[int index] { get; }

    /// <summary>
    /// 根据名称获取效果参数。
    /// </summary>
    IOfficeEffectParameter this[string name] { get; }

    /// <summary>
    /// 检查是否包含指定名称的参数。
    /// </summary>
    /// <param name="name">参数名称。</param>
    /// <returns>是否包含。</returns>
    bool Contains(string name);

    /// <summary>
    /// 获取所有参数名称。
    /// </summary>
    /// <returns>参数名称列表。</returns>
    List<string> GetAllParameterNames();

    /// <summary>
    /// 设置参数值。
    /// </summary>
    /// <param name="name">参数名称。</param>
    /// <param name="value">参数值。</param>
    /// <returns>是否设置成功。</returns>
    bool SetValue(string name, object value);

    /// <summary>
    /// 获取参数值。
    /// </summary>
    /// <param name="name">参数名称。</param>
    /// <returns>参数值。</returns>
    object GetValue(string name);

    /// <summary>
    /// 复制参数设置到另一个参数集合。
    /// </summary>
    /// <param name="targetParameters">目标参数集合。</param>
    void CopyTo(IOfficeEffectParameters targetParameters);
}
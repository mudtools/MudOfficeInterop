namespace MudTools.OfficeInterop;

/// <summary>
/// 封装 Microsoft.Office.Core.EffectParameter 的接口，用于操作效果参数。
/// </summary>
public interface IOfficeEffectParameter : IDisposable
{
    /// <summary>
    /// 获取参数的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置参数的值。
    /// </summary>
    object Value { get; set; }

    /// <summary>
    /// 获取参数值的字符串表示。
    /// </summary>
    /// <returns>参数值的字符串表示。</returns>
    string GetValueAsString();
}
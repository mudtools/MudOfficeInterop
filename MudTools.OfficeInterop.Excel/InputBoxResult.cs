//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 封装 Excel InputBox 操作的结果。
/// </summary>
public class InputBoxResult<T>
{
    /// <summary>
    /// 获取操作的结果类型。
    /// </summary>
    public InputBoxResultType ResultType { get; }

    /// <summary>
    /// 获取用户输入的值。如果 ResultType 不是 Ok，则此值可能为 null 或无意义。
    /// 返回类型取决于调用 InputBox 时指定的 Type 参数。
    /// </summary>
    public T? Value { get; internal set; }

    /// <summary>
    /// 初始化 InputBoxResult 类的新实例。
    /// </summary>
    /// <param name="type">结果类型。</param>
    /// <param name="value">用户输入的值。</param>
    public InputBoxResult(InputBoxResultType type, T? value)
    {
        ResultType = type;
        Value = value;
    }
}

/// <summary>
/// 封装 Excel InputBox 操作的结果。
/// </summary>
public class InputBoxResult : InputBoxResult<object>
{
    /// <summary>
    /// 初始化 InputBoxResult 类的新实例。
    /// </summary>
    /// <param name="type">结果类型。</param>
    /// <param name="value">用户输入的值。</param>
    public InputBoxResult(InputBoxResultType type, object? value)
        : base(type, value)
    {
    }
}

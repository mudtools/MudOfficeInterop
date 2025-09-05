//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

internal static class StringEx
{
    static StringEx()
        => _columnNumberStorage = new Dictionary<string, int>();

    public static int ToColumnNumber(this string @this)
    {
        if (!@this.IsValidColumnName())
            throw new ArgumentOutOfRangeException("The string must contain only A to Z.");

        if (_columnNumberStorage.TryGetValue(@this, out int n))
            return n;

        var index = 0;
        foreach (int col in @this.ToUpper())
            index = (index * 26) + (col - 'A' + 1);

        _columnNumberStorage.Add(@this, index);
        return index;
    }

    private static bool IsValidColumnName(this string @this)
    {
        foreach (uint c in @this.ToUpper())
        {
            if (c < 'A' && 'Z' < c)
                return false;
        }
        return true;
    }

    public static string Replace(this string text, string oldValue, string newValue, StringComparison comparison, ref int count)
    {
        if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(oldValue))
            return text;

        var result = text;
        var index = result.IndexOf(oldValue, comparison);
        count = 0;

        while (index >= 0)
        {
            result = result.Substring(0, index) + newValue + result.Substring(index + oldValue.Length);
            count++;
            index = result.IndexOf(oldValue, index + newValue.Length, comparison);
        }

        return result;
    }

    public static object ComArgsVal(this string? val, Func<string, bool>? condition = null)
    {
        if (val != null && !string.IsNullOrEmpty(val))
        {
            return val;
        }
        return Type.Missing;
    }

    private static readonly Dictionary<string, int> _columnNumberStorage;
}

//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
internal static class ObjectEx
{
    public static object ComArgsVal<T>(this Nullable<T> val, Func<T, bool>? condition = null)
        where T : struct
    {
        if (val != null && val.HasValue)
        {
            if (condition == null)
                return val.Value;
            if (condition(val.Value))
                return val.Value;
        }
        return Type.Missing;
    }

    public static bool IsExcelDateSerial(this double value)
    {
        // Excel日期范围：1900年1月1日到9999年12月31日
        return value >= 1 && value <= 2958465.99999;
    }

    public static double ConvertToDouble(this object result)
    {
        if (result is double d) return d;
        if (result is int i) return i;
        if (result is float f) return f;
        if (result is decimal dec) return (double)dec;

        if (double.TryParse(result?.ToString(), out double parsed))
            return parsed;

        throw new InvalidCastException($"Cannot convert {result?.GetType().Name ?? "null"} to double.");
    }

    public static float ConvertToFloat(this object result)
    {
        if (result is double d) return (float)d;
        if (result is int i) return i;
        if (result is float f) return f;
        if (result is decimal dec) return (float)dec;

        if (float.TryParse(result?.ToString(), out float parsed))
            return parsed;

        throw new InvalidCastException($"Cannot convert {result?.GetType().Name ?? "null"} to double.");
    }

    public static bool ConvertToBool(this object result)
    {
        if (result is bool b) return b;
        if (result is int i) return i != 0;
        if (result is double d) return d != 0;

        if (bool.TryParse(result?.ToString(), out bool parsed))
            return parsed;

        throw new InvalidCastException($"Cannot convert {result?.GetType().Name ?? "null"} to bool.");
    }

    public static DateTime ConvertToDateTime(this object result)
    {
        if (result is DateTime dt) return dt;
        if (result is double d && IsExcelDateSerial(d))
            return DateTime.FromOADate(d);

        if (DateTime.TryParse(result?.ToString(), out DateTime parsed))
            return parsed;

        throw new InvalidCastException($"Cannot convert {result?.GetType().Name ?? "null"} to DateTime.");
    }

    public static object[,] ConvertToArray(this object result)
    {
        if (result is object[,] array) return array;

        // 单值转换为1x1数组
        object[,] singleValueArray = new object[1, 1];
        singleValueArray[0, 0] = result;
        return singleValueArray;
    }
}

//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

internal static class ObjectEx
{

    /// <summary>
    /// 实现枚举值之间的安全转换
    /// </summary>
    /// <typeparam name="T">源枚举类型</typeparam>
    /// <typeparam name="TReturn">目标枚举类型</typeparam>
    /// <param name="defaultVal">转换失败时返回的默认值。</param>
    /// <param name="val">源枚举值</param>
    /// <returns>目标枚举值</returns>
    public static TReturn EnumConvert<T, TReturn>(this T val, TReturn defaultVal = default)
    where T : struct, Enum
    where TReturn : struct, Enum
    {
        // 获取源枚举值的底层数值
        var underlyingValue = Convert.ChangeType(val, Enum.GetUnderlyingType(typeof(T)));

        // 检查该值是否在目标枚举的范围内
        if (Enum.IsDefined(typeof(TReturn), underlyingValue))
        {
            return (TReturn)Enum.ToObject(typeof(TReturn), underlyingValue);
        }

        // 可选：处理值不在目标枚举中的情况
        // 这里返回 null，但您可能希望抛出异常或使用其他处理方式
        return defaultVal;
    }

    /// <summary>
    /// 实现枚举值之间的安全转换
    /// </summary>
    /// <typeparam name="T">源枚举类型</typeparam>
    /// <typeparam name="TReturn">目标枚举类型</typeparam>
    /// <param name="defaultVal">转换失败时返回的默认值。</param>
    /// <param name="val">源枚举值</param>
    /// <returns>目标枚举值</returns>
    public static TReturn EnumConvert<T, TReturn>(this T? val, TReturn defaultVal = default)
    where T : struct, Enum
    where TReturn : struct, Enum
    {
        if (!val.HasValue)
            return defaultVal;

        // 获取源枚举值的底层数值
        var underlyingValue = Convert.ChangeType(val.Value, Enum.GetUnderlyingType(typeof(T)));

        // 检查该值是否在目标枚举的范围内
        if (Enum.IsDefined(typeof(TReturn), underlyingValue))
        {
            return (TReturn)Enum.ToObject(typeof(TReturn), underlyingValue);
        }

        // 可选：处理值不在目标枚举中的情况
        // 这里返回 null，但您可能希望抛出异常或使用其他处理方式
        return defaultVal;
    }

    /// <summary>
    /// 实现枚举值之间的安全转换
    /// </summary>
    /// <typeparam name="T">源枚举类型</typeparam>
    /// <typeparam name="TReturn">目标枚举类型</typeparam>
    /// <param name="defaultVal">转换失败时返回的默认值。</param>
    /// <param name="val">源枚举值</param>
    /// <returns>目标枚举值</returns>
    public static TReturn? EnumConvert<T, TReturn>(this T? val, TReturn? defaultVal = default)
    where T : struct, Enum
    where TReturn : struct, Enum
    {
        if (!val.HasValue)
            return defaultVal;

        // 获取源枚举值的底层数值
        var underlyingValue = Convert.ChangeType(val.Value, Enum.GetUnderlyingType(typeof(T)));

        // 检查该值是否在目标枚举的范围内
        if (Enum.IsDefined(typeof(TReturn), underlyingValue))
        {
            return (TReturn)Enum.ToObject(typeof(TReturn), underlyingValue);
        }

        // 可选：处理值不在目标枚举中的情况
        // 这里返回 null，但您可能希望抛出异常或使用其他处理方式
        return defaultVal;
    }

    /// <summary>
    /// 将可空值转换为 COM 参数值，若为空或不满足条件则返回 Type.Missing。
    /// </summary>
    public static object ComArgsVal<T>(this T? val, Func<T, bool>? condition = null)
        where T : struct
    {
        if (val.HasValue && (condition == null || condition(val.Value)))
            return val.Value;
        return Type.Missing;
    }

    /// <summary>
    /// 将可空值转换为 COM 参数值，若为空或不满足条件则返回空值，否则返回转换值。
    /// </summary>
    public static TReturn? ComArgsConvert<T, TReturn>(this T? val, Func<T, TReturn>? convert)
         where T : struct
        where TReturn : struct
    {
        if (!val.HasValue || convert == null)
            return default;
        return convert(val.Value);
    }

    /// <summary>
    /// 判断双精度值是否为 Excel 日期序列号（1900-01-01 至 9999-12-31）
    /// </summary>
    public static bool IsExcelDateSerial(this double value)
    {
        return value >= 1 && value <= 2958465.99999;
    }

    /// <summary>
    /// 将对象转换为 double
    /// </summary>
    public static double ConvertToDouble(this object result)
    {
        return result switch
        {
            double d => d,
            int i => i,
            float f => f,
            long l => l,
            short s => s,
            decimal dec => (double)dec,
            _ => TryParseOrThrow<double>(result, double.TryParse, nameof(Double))
        };
    }

    /// <summary>
    /// 将对象转换为 float
    /// </summary>
    public static float ConvertToFloat(this object result)
    {
        return result switch
        {
            double d => (float)d,
            int i => i,
            float f => f,
            long l => l,
            short s => s,
            decimal dec => (float)dec,
            _ => TryParseOrThrow<float>(result, float.TryParse, nameof(Single))
        };
    }

    /// <summary>
    /// 将对象转换为 bool
    /// </summary>
    public static bool ConvertToBool(this object result)
    {
        return result switch
        {
            bool b => b,
            int i => i != 0,
            short s => s != 0,
            long l => l != 0,
            double d => d != 0,
            _ => TryParseOrThrow<bool>(result, bool.TryParse, nameof(Boolean))
        };
    }

    /// <summary>
    /// 将对象转换为 DateTime
    /// </summary>
    public static DateTime ConvertToDateTime(this object result)
    {
        if (result is DateTime dt) return dt;
        if (result is double d && d.IsExcelDateSerial())
            return DateTime.FromOADate(d);

        return TryParseOrThrow<DateTime>(result, DateTime.TryParse, nameof(DateTime));
    }

    /// <summary>
    /// 将对象转换为二维对象数组（单值转为 1x1 数组）
    /// </summary>
    public static object[,] ConvertToArray(this object result)
    {
        return result is object[,] array ? array : new object[1, 1] { { result } };
    }

    // ========== 私有辅助方法 ==========

    private static T TryParseOrThrow<T>(object result, TryParseDelegate<T> tryParse, string targetType)
    {
        if (result == null)
            throw new InvalidCastException($"Cannot convert null to {targetType}.");

        if (tryParse(result.ToString(), out T parsed))
            return parsed;

        throw new InvalidCastException($"Cannot convert {result.GetType().Name} to {targetType}.");
    }

    private delegate bool TryParseDelegate<T>(string s, out T result);
}
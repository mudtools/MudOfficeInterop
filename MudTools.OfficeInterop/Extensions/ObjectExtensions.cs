//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Collections.Concurrent;
using System.Linq.Expressions;

namespace MudTools.OfficeInterop;

internal static class ObjectExtensions
{
    // 缓存无参构造函数
    private static readonly ConcurrentDictionary<Type, Func<object>?> _constructorCache = new();

    /// <summary>
    /// 创建Office UI对象的包装器实例
    /// 此方法通过反射查找与接口T对应的实现类，并将COM对象包装为强类型的接口实例
    /// </summary>
    /// <typeparam name="T">Office对象接口类型，必须实现IOfficeObject&lt;T&gt;接口</typeparam>
    /// <param name="comObj">原始的COM对象，将被包装为接口T的实例</param>
    /// <returns>接口T的实现实例，如果无法创建则返回默认值(null或类型的默认值)</returns>
    public static T? Create<T>(object comObj) where T : IOfficeObject<T>
    {
        var type = typeof(T);
        var impFullName = GetImpTypeName(type);
        // 查找实现类，该类必须是class类型且全名匹配

        var types = type.Assembly.GetTypes().Where(t => t.IsClass && !t.IsAbstract).ToList();
        var implementationType = types
                        .Where(t => t.IsClass && !t.IsAbstract
                                    && t.FullName.Equals(impFullName, StringComparison.Ordinal))
                        .FirstOrDefault();

        if (implementationType == null) return default;

        // 通过反射创建实现类的实例
        var instance = CreateInstance(implementationType);
        if (instance == null)
            return default;
        // 将实例转换为接口T，并使用传入的COM对象加载数据
        if (instance is T t)
            return t.LoadFromObject(comObj);
        return default;
    }

    private static string GetImpTypeName(Type type)
    {
        return $"{type.Namespace}.Imps.{type.Name.TrimStart('I')}";
    }

    private static object? CreateInstance(Type type)
    {
        return _constructorCache.GetOrAdd(type, t =>
        {
            try
            {
                var constructor = t.GetConstructor(Type.EmptyTypes);
                if (constructor == null)
                    return null;

                return Expression.Lambda<Func<object>>(
                            Expression.New(constructor))
                            .Compile();
            }
            catch
            {
                return null;
            }
        })?.Invoke();
    }


    /// <summary>
    /// 将枚举值转换为另一种枚举类型（非空到非空）
    /// </summary>
    public static TReturn EnumConvert<TReturn>(this object val, TReturn defaultVal = default)
        where TReturn : struct, Enum
    {
        return ConvertEnumValue(val, defaultVal);
    }

    /// <summary>
    /// 将枚举值转换为另一种枚举类型（非空到非空）
    /// </summary>
    public static TReturn EnumConvert<T, TReturn>(this T val, TReturn defaultVal = default)
        where T : struct, Enum
        where TReturn : struct, Enum
    {
        return ConvertEnumValue(val, defaultVal);
    }

    /// <summary>
    /// 将可空枚举值转换为另一种枚举类型（可空到非空）
    /// </summary>
    public static TReturn EnumConvert<T, TReturn>(this T? val, TReturn defaultVal = default)
        where T : struct, Enum
        where TReturn : struct, Enum
    {
        return val.HasValue ? ConvertEnumValue(val.Value, defaultVal) : defaultVal;
    }

    /// <summary>
    /// 将可空枚举值转换为另一种可空枚举类型（可空到可空）
    /// </summary>
    public static TReturn? EnumConvert<T, TReturn>(this T? val, TReturn? defaultVal = default)
        where T : struct, Enum
        where TReturn : struct, Enum
    {
        if (!val.HasValue)
            return defaultVal;

        var underlyingValue = GetUnderlyingValue(val.Value);
        return Enum.IsDefined(typeof(TReturn), underlyingValue)
            ? (TReturn)Enum.ToObject(typeof(TReturn), underlyingValue)
            : defaultVal;
    }

    /// <summary>
    /// 内部转换方法，处理核心转换逻辑
    /// </summary>
    private static TReturn ConvertEnumValue<T, TReturn>(T val, TReturn defaultVal)
        where T : struct, Enum
        where TReturn : struct, Enum
    {
        var underlyingValue = GetUnderlyingValue(val);
        return Enum.IsDefined(typeof(TReturn), underlyingValue)
            ? (TReturn)Enum.ToObject(typeof(TReturn), underlyingValue)
            : defaultVal;
    }

    private static TReturn ConvertEnumValue<TReturn>(object val, TReturn defaultVal)
    where TReturn : struct, Enum
    {
        var targetType = typeof(TReturn);
        var underlyingType = Enum.GetUnderlyingType(targetType);

        object converted;
        try
        {
            converted = Convert.ChangeType(val, underlyingType);
        }
        catch (Exception)
        {
            return defaultVal;
        }

        return Enum.IsDefined(targetType, converted)
            ? (TReturn)Enum.ToObject(targetType, converted)
            : defaultVal;
    }


    /// <summary>
    /// 获取枚举的底层数值
    /// </summary>
    private static object GetUnderlyingValue<T>(T val) where T : struct, Enum
    {
        return Convert.ChangeType(val, Enum.GetUnderlyingType(typeof(T)));
    }

    /// <summary>
    /// 将对象转换为指定的枚举类型
    /// </summary>
    /// <typeparam name="TReturn">目标枚举类型</typeparam>
    /// <param name="val">需要转换的对象值</param>
    /// <param name="defaultVal">当转换失败时返回的默认值</param>
    /// <returns>转换成功则返回对应的枚举值，否则返回默认值</returns>
    public static TReturn ObjectConvertEnum<TReturn>(this object? val, TReturn defaultVal = default)
    where TReturn : struct, Enum
    {
        if (val == null)
            return defaultVal;

        try
        {
            var targetType = typeof(TReturn);
            var underlyingType = Enum.GetUnderlyingType(targetType);

            // 检查val是否已经是目标枚举类型
            if (val.GetType() == targetType)
                return (TReturn)val;

            var underlyingValue = Convert.ChangeType(val, underlyingType);
            return Enum.IsDefined(targetType, underlyingValue)
                ? (TReturn)Enum.ToObject(targetType, underlyingValue)
                : defaultVal;
        }
        catch (Exception) when (val != null)
        {
            return defaultVal;
        }
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
    /// 将对象转换为 decimal
    /// </summary>
    public static decimal ConvertToDecimal(this object result)
    {
        return result switch
        {
            double d => (decimal)d,
            int i => i,
            float f => (decimal)f,
            long l => l,
            short s => s,
            decimal dec => dec,
            _ => TryParseOrThrow<decimal>(result, decimal.TryParse, nameof(Decimal))
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
    /// 将对象转换为 Int
    /// </summary>
    public static int ConvertToInt(this object result)
    {
        return result switch
        {
            double d => (int)d,
            int i => i,
            float f => (int)f,
            long l => (int)l,
            short s => s,
            decimal dec => (int)dec,
            _ => TryParseOrThrow<int>(result, int.TryParse, nameof(Int32))
        };
    }

    /// <summary>
    /// 将对象转换为 long
    /// </summary>
    public static long ConvertToLong(this object result)
    {
        return result switch
        {
            double d => (long)d,
            int i => i,
            float f => (long)f,
            long l => l,
            short s => s,
            decimal dec => (long)dec,
            _ => TryParseOrThrow<long>(result, long.TryParse, nameof(Int64))
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
            MsCore.MsoTriState t => t == MsCore.MsoTriState.msoTrue || t == MsCore.MsoTriState.msoCTrue,
            _ => TryParseOrThrow<bool>(result, bool.TryParse, nameof(Boolean))
        };
    }

    public static MsCore.MsoTriState ConvertTriState(this bool? b)
    {
        return b != null && b.Value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
    }

    public static MsCore.MsoTriState ConvertTriState(this bool b)
    {
        return b ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
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
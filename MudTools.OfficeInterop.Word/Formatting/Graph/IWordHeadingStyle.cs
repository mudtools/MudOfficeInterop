
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示用于构建目录或图表目录的单个标题样式的二次封装接口。
/// 此接口封装了样式本身及其在目录中的层级信息。
/// </summary>
public interface IWordHeadingStyle : IDisposable
{
    /// <summary>
    /// 获取此标题样式所属的 Word 应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此标题样式的父对象（通常是 <see cref="IWordHeadingStyles"/> 集合）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置此标题样式在目录中的层级。
    /// 有效值为 1 到 9，对应于目录中的标题级别。
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">当设置的值不在 1-9 范围内时可能抛出（由底层 COM 对象决定）。</exception>
    int Level { get; set; }

    /// <summary>
    /// 获取或设置此标题样式所关联的 Word 样式对象。
    /// 可以是内置样式（如 WdBuiltinStyle.wdStyleHeading1）或自定义样式。
    /// </summary>
    /// <remarks>
    /// 根据官方文档，可以为此属性指定样式的本地名称、整数、WdBuiltinStyle 常量或表示样式的对象 [[6]]。
    /// 在此封装中，我们使用 <see cref="IWordStyle"/> 接口来提供强类型访问。
    /// </remarks>
    IWordStyle? Style { get; set; }

    void Delete();
}
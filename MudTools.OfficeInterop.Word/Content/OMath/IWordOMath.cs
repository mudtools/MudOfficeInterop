//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中的数学对象，提供对Word文档中数学公式的访问和操作功能
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMath : IDisposable
{
    /// <summary>
    /// 获取与该数学对象关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取该数学对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取与该数学对象关联的文本范围
    /// </summary>
    IWordRange Range { get; }

    /// <summary>
    /// 获取该数学对象的父级数学对象
    /// </summary>
    IWordOMath ParentOMath { get; }

    /// <summary>
    /// 获取与该数学对象关联的数学函数集合
    /// </summary>
    IWordOMathFunctions Functions { get; }

    /// <summary>
    /// 获取该数学对象所属的父级数学函数
    /// </summary>
    IWordOMathFunction ParentFunction { get; }

    /// <summary>
    /// 获取该数学对象所属的矩阵行
    /// </summary>
    IWordOMathMatRow ParentRow { get; }

    /// <summary>
    /// 获取该数学对象所属的矩阵列
    /// </summary>
    IWordOMathMatCol ParentCol { get; }

    /// <summary>
    /// 获取该数学对象中的换行符集合
    /// </summary>
    IWordOMathBreaks Breaks { get; }

    /// <summary>
    /// 获取该数学对象的父级参数对象
    /// </summary>
    IWordOMath ParentArg { get; }

    /// <summary>
    /// 获取参数索引
    /// </summary>
    int ArgIndex { get; }

    /// <summary>
    /// 获取嵌套级别
    /// </summary>
    int NestingLevel { get; }

    /// <summary>
    /// 获取或设置参数大小
    /// </summary>
    int ArgSize { get; set; }

    /// <summary>
    /// 获取或设置数学对象的类型
    /// </summary>
    WdOMathType Type { get; set; }

    /// <summary>
    /// 获取或设置数学对象的对齐方式
    /// </summary>
    WdOMathJc Justification { get; set; }

    /// <summary>
    /// 获取或设置对齐点
    /// </summary>
    int AlignPoint { get; set; }

    /// <summary>
    /// 将数学对象线性化显示
    /// </summary>
    void Linearize();

    /// <summary>
    /// 构建数学对象的格式化显示
    /// </summary>
    void BuildUp();

    /// <summary>
    /// 移除该数学对象
    /// </summary>
    void Remove();

    /// <summary>
    /// 将数学对象转换为数学文本格式
    /// </summary>
    void ConvertToMathText();

    /// <summary>
    /// 将数学对象转换为普通文本格式
    /// </summary>
    void ConvertToNormalText();

    /// <summary>
    /// 将数学对象转换为字面文本格式
    /// </summary>
    void ConvertToLiteralText();
}
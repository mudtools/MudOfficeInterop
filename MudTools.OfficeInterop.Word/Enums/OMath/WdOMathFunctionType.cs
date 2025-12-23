//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定公式函数的类型
/// </summary>
[Guid("74779721-3C00-363D-BED4-B0AF3595EB05")]
public enum WdOMathFunctionType
{
    /// <summary>
    /// 公式重音标记
    /// </summary>
    wdOMathFunctionAcc = 1,

    /// <summary>
    /// 公式分数线
    /// </summary>
    wdOMathFunctionBar,

    /// <summary>
    /// 方框
    /// </summary>
    wdOMathFunctionBox,

    /// <summary>
    /// 边框方框
    /// </summary>
    wdOMathFunctionBorderBox,

    /// <summary>
    /// 公式分隔符
    /// </summary>
    wdOMathFunctionDelim,

    /// <summary>
    /// 公式数组
    /// </summary>
    wdOMathFunctionEqArray,

    /// <summary>
    /// 公式分数
    /// </summary>
    wdOMathFunctionFrac,

    /// <summary>
    /// 公式函数
    /// </summary>
    wdOMathFunctionFunc,

    /// <summary>
    /// 分组字符
    /// </summary>
    wdOMathFunctionGroupChar,

    /// <summary>
    /// 公式下限
    /// </summary>
    wdOMathFunctionLimLow,

    /// <summary>
    /// 公式上限
    /// </summary>
    wdOMathFunctionLimUpp,

    /// <summary>
    /// 公式矩阵
    /// </summary>
    wdOMathFunctionMat,

    /// <summary>
    /// 公式 n 元运算符
    /// </summary>
    wdOMathFunctionNary,

    /// <summary>
    /// 公式幻影
    /// </summary>
    wdOMathFunctionPhantom,

    /// <summary>
    /// 预置脚本
    /// </summary>
    wdOMathFunctionScrPre,

    /// <summary>
    /// 公式根式
    /// </summary>
    wdOMathFunctionRad,

    /// <summary>
    /// 下标脚本
    /// </summary>
    wdOMathFunctionScrSub,

    /// <summary>
    /// 上下标脚本
    /// </summary>
    wdOMathFunctionScrSubSup,

    /// <summary>
    /// 上标脚本
    /// </summary>
    wdOMathFunctionScrSup,

    /// <summary>
    /// 公式文本
    /// </summary>
    wdOMathFunctionText,

    /// <summary>
    /// 公式常规文本
    /// </summary>
    wdOMathFunctionNormalText,

    /// <summary>
    /// 公式字面文本
    /// </summary>
    wdOMathFunctionLiteralText
}
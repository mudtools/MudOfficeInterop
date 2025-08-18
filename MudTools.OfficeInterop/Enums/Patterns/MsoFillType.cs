//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定填充类型的枚举，用于形状或对象的填充样式
/// </summary>
public enum MsoFillType
{
    /// <summary>
    /// 混合填充类型
    /// </summary>
    msoFillMixed = -2,

    /// <summary>
    /// 纯色填充
    /// </summary>
    msoFillSolid = 1,

    /// <summary>
    /// 图案填充
    /// </summary>
    msoFillPatterned = 2,

    /// <summary>
    /// 渐变填充
    /// </summary>
    msoFillGradient = 3,

    /// <summary>
    /// 纹理填充
    /// </summary>
    msoFillTextured = 4,

    /// <summary>
    /// 背景填充
    /// </summary>
    msoFillBackground = 5,

    /// <summary>
    /// 图片填充
    /// </summary>
    msoFillPicture = 6
}
//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定纹理在对象中的对齐方式
/// </summary>
public enum MsoTextureAlignment
{
    /// <summary>
    /// 混合纹理对齐方式
    /// </summary>
    msoTextureAlignmentMixed = -2,

    /// <summary>
    /// 纹理在对象顶部左对齐
    /// </summary>
    msoTextureTopLeft = 0,

    /// <summary>
    /// 纹理在对象顶部居中对齐
    /// </summary>
    msoTextureTop = 1,

    /// <summary>
    /// 纹理在对象顶部右对齐
    /// </summary>
    msoTextureTopRight = 2,

    /// <summary>
    /// 纹理在对象左侧居中对齐
    /// </summary>
    msoTextureLeft = 3,

    /// <summary>
    /// 纹理在对象中心对齐
    /// </summary>
    msoTextureCenter = 4,

    /// <summary>
    /// 纹理在对象右侧居中对齐
    /// </summary>
    msoTextureRight = 5,

    /// <summary>
    /// 纹理在对象底部左对齐
    /// </summary>
    msoTextureBottomLeft = 6,

    /// <summary>
    /// 纹理在对象底部居中对齐
    /// </summary>
    msoTextureBottom = 7,

    /// <summary>
    /// 纹理在对象底部右对齐
    /// </summary>
    msoTextureBottomRight = 8
}
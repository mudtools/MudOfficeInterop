namespace MudTools.OfficeInterop;

/// <summary>
/// 指定颜色类型，用于Office互操作中的颜色定义
/// </summary>
public enum MsoColorType
{
    /// <summary>
    /// 混合颜色类型，表示多种颜色类型的组合
    /// </summary>
    msoColorTypeMixed = -2,

    /// <summary>
    /// RGB颜色类型，基于红绿蓝三原色的颜色模型
    /// </summary>
    msoColorTypeRGB = 1,

    /// <summary>
    /// 方案颜色类型，基于Office主题方案的颜色
    /// </summary>
    msoColorTypeScheme = 2,

    /// <summary>
    /// CMYK颜色类型，基于青色、品红、黄色和黑色的颜色模型
    /// </summary>
    msoColorTypeCMYK = 3,

    /// <summary>
    /// CMS颜色类型，基于颜色管理系统定义的颜色
    /// </summary>
    msoColorTypeCMS = 4,

    /// <summary>
    /// 墨水颜色类型，通常用于印刷相关的颜色定义
    /// </summary>
    msoColorTypeInk = 5
}
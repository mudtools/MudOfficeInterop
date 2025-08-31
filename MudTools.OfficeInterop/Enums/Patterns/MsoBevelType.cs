namespace MudTools.OfficeInterop;

/// <summary>
/// 指定在Office应用程序中使用的斜面类型（Bevel Type）
/// </summary>
public enum MsoBevelType
{
    /// <summary>
    /// 混合斜面类型
    /// </summary>
    msoBevelTypeMixed = -2,
    /// <summary>
    /// 无斜面效果
    /// </summary>
    msoBevelNone = 1,
    /// <summary>
    /// 宽松插入斜面效果
    /// </summary>
    msoBevelRelaxedInset = 2,
    /// <summary>
    /// 圆形斜面效果
    /// </summary>
    msoBevelCircle = 3,
    /// <summary>
    /// 斜坡斜面效果
    /// </summary>
    msoBevelSlope = 4,
    /// <summary>
    /// 交叉斜面效果
    /// </summary>
    msoBevelCross = 5,
    /// <summary>
    /// 角度斜面效果
    /// </summary>
    msoBevelAngle = 6,
    /// <summary>
    /// 软圆形斜面效果
    /// </summary>
    msoBevelSoftRound = 7,
    /// <summary>
    /// 凸面斜面效果
    /// </summary>
    msoBevelConvex = 8,
    /// <summary>
    /// 冷斜面效果
    /// </summary>
    msoBevelCoolSlant = 9,
    /// <summary>
    /// 麻点斜面效果
    /// </summary>
    msoBevelDivot = 10,
    /// <summary>
    /// 棱面斜面效果
    /// </summary>
    msoBevelRiblet = 11,
    /// <summary>
    /// 硬边斜面效果
    /// </summary>
    msoBevelHardEdge = 12,
    /// <summary>
    /// 艺术装饰斜面效果
    /// </summary>
    msoBevelArtDeco = 13
}
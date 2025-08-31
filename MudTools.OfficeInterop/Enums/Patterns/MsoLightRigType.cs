namespace MudTools.OfficeInterop;

/// <summary>
/// 指定在Office应用程序中应用于形状的照明设置类型
/// </summary>
public enum MsoLightRigType
{
    /// <summary>
    /// 混合照明类型
    /// </summary>
    msoLightRigMixed = -2,
    /// <summary>
    /// 传统平面照明1
    /// </summary>
    msoLightRigLegacyFlat1 = 1,
    /// <summary>
    /// 传统平面照明2
    /// </summary>
    msoLightRigLegacyFlat2 = 2,
    /// <summary>
    /// 传统平面照明3
    /// </summary>
    msoLightRigLegacyFlat3 = 3,
    /// <summary>
    /// 传统平面照明4
    /// </summary>
    msoLightRigLegacyFlat4 = 4,
    /// <summary>
    /// 传统普通照明1
    /// </summary>
    msoLightRigLegacyNormal1 = 5,
    /// <summary>
    /// 传统普通照明2
    /// </summary>
    msoLightRigLegacyNormal2 = 6,
    /// <summary>
    /// 传统普通照明3
    /// </summary>
    msoLightRigLegacyNormal3 = 7,
    /// <summary>
    /// 传统普通照明4
    /// </summary>
    msoLightRigLegacyNormal4 = 8,
    /// <summary>
    /// 传统强烈照明1
    /// </summary>
    msoLightRigLegacyHarsh1 = 9,
    /// <summary>
    /// 传统强烈照明2
    /// </summary>
    msoLightRigLegacyHarsh2 = 10,
    /// <summary>
    /// 传统强烈照明3
    /// </summary>
    msoLightRigLegacyHarsh3 = 11,
    /// <summary>
    /// 传统强烈照明4
    /// </summary>
    msoLightRigLegacyHarsh4 = 12,
    /// <summary>
    /// 三点照明
    /// </summary>
    msoLightRigThreePoint = 13,
    /// <summary>
    /// 平衡照明
    /// </summary>
    msoLightRigBalanced = 14,
    /// <summary>
    /// 柔和照明
    /// </summary>
    msoLightRigSoft = 15,
    /// <summary>
    /// 强烈照明
    /// </summary>
    msoLightRigHarsh = 16,
    /// <summary>
    /// 泛光照明
    /// </summary>
    msoLightRigFlood = 17,
    /// <summary>
    /// 对比照明
    /// </summary>
    msoLightRigContrasting = 18,
    /// <summary>
    /// 晨光照明
    /// </summary>
    msoLightRigMorning = 19,
    /// <summary>
    /// 日出照明
    /// </summary>
    msoLightRigSunrise = 20,
    /// <summary>
    /// 日落照明
    /// </summary>
    msoLightRigSunset = 21,
    /// <summary>
    /// 寒冷照明
    /// </summary>
    msoLightRigChilly = 22,
    /// <summary>
    /// 冰冻照明
    /// </summary>
    msoLightRigFreezing = 23,
    /// <summary>
    /// 平面照明
    /// </summary>
    msoLightRigFlat = 24,
    /// <summary>
    /// 双点照明
    /// </summary>
    msoLightRigTwoPoint = 25,
    /// <summary>
    /// 发光照明
    /// </summary>
    msoLightRigGlow = 26,
    /// <summary>
    /// 明亮房间照明
    /// </summary>
    msoLightRigBrightRoom = 27
}
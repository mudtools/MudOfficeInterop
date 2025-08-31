namespace MudTools.OfficeInterop;

/// <summary>
/// 指定用于三维形状的预设摄像机视图
/// </summary>
public enum MsoPresetCamera
{
    /// <summary>
    /// 混合摄像机设置
    /// </summary>
    msoPresetCameraMixed = -2,
    /// <summary>
    /// 左上角的传统斜角视图
    /// </summary>
    msoCameraLegacyObliqueTopLeft = 1,
    /// <summary>
    /// 顶部的传统斜角视图
    /// </summary>
    msoCameraLegacyObliqueTop = 2,
    /// <summary>
    /// 右上角的传统斜角视图
    /// </summary>
    msoCameraLegacyObliqueTopRight = 3,
    /// <summary>
    /// 左侧的传统斜角视图
    /// </summary>
    msoCameraLegacyObliqueLeft = 4,
    /// <summary>
    /// 前面的传统斜角视图
    /// </summary>
    msoCameraLegacyObliqueFront = 5,
    /// <summary>
    /// 右侧的传统斜角视图
    /// </summary>
    msoCameraLegacyObliqueRight = 6,
    /// <summary>
    /// 左下角的传统斜角视图
    /// </summary>
    msoCameraLegacyObliqueBottomLeft = 7,
    /// <summary>
    /// 底部的传统斜角视图
    /// </summary>
    msoCameraLegacyObliqueBottom = 8,
    /// <summary>
    /// 右下角的传统斜角视图
    /// </summary>
    msoCameraLegacyObliqueBottomRight = 9,
    /// <summary>
    /// 左上角的传统透视视图
    /// </summary>
    msoCameraLegacyPerspectiveTopLeft = 10,
    /// <summary>
    /// 顶部的传统透视视图
    /// </summary>
    msoCameraLegacyPerspectiveTop = 11,
    /// <summary>
    /// 右上角的传统透视视图
    /// </summary>
    msoCameraLegacyPerspectiveTopRight = 12,
    /// <summary>
    /// 左侧的传统透视视图
    /// </summary>
    msoCameraLegacyPerspectiveLeft = 13,
    /// <summary>
    /// 前面的传统透视视图
    /// </summary>
    msoCameraLegacyPerspectiveFront = 14,
    /// <summary>
    /// 右侧的传统透视视图
    /// </summary>
    msoCameraLegacyPerspectiveRight = 15,
    /// <summary>
    /// 左下角的传统透视视图
    /// </summary>
    msoCameraLegacyPerspectiveBottomLeft = 16,
    /// <summary>
    /// 底部的传统透视视图
    /// </summary>
    msoCameraLegacyPerspectiveBottom = 17,
    /// <summary>
    /// 右下角的传统透视视图
    /// </summary>
    msoCameraLegacyPerspectiveBottomRight = 18,
    /// <summary>
    /// 正面正交视图
    /// </summary>
    msoCameraOrthographicFront = 19,
    /// <summary>
    /// 向上等轴测顶视图
    /// </summary>
    msoCameraIsometricTopUp = 20,
    /// <summary>
    /// 向下等轴测顶视图
    /// </summary>
    msoCameraIsometricTopDown = 21,
    /// <summary>
    /// 向上等轴测底视图
    /// </summary>
    msoCameraIsometricBottomUp = 22,
    /// <summary>
    /// 向下等轴测底视图
    /// </summary>
    msoCameraIsometricBottomDown = 23,
    /// <summary>
    /// 向上等轴测左视图
    /// </summary>
    msoCameraIsometricLeftUp = 24,
    /// <summary>
    /// 向下等轴测左视图
    /// </summary>
    msoCameraIsometricLeftDown = 25,
    /// <summary>
    /// 向上等轴测右视图
    /// </summary>
    msoCameraIsometricRightUp = 26,
    /// <summary>
    /// 向下等轴测右视图
    /// </summary>
    msoCameraIsometricRightDown = 27,
    /// <summary>
    /// 第一轴向左等轴测视图
    /// </summary>
    msoCameraIsometricOffAxis1Left = 28,
    /// <summary>
    /// 第一轴向右等轴测视图
    /// </summary>
    msoCameraIsometricOffAxis1Right = 29,
    /// <summary>
    /// 第一轴向上等轴测视图
    /// </summary>
    msoCameraIsometricOffAxis1Top = 30,
    /// <summary>
    /// 第二轴向左等轴测视图
    /// </summary>
    msoCameraIsometricOffAxis2Left = 31,
    /// <summary>
    /// 第二轴向右等轴测视图
    /// </summary>
    msoCameraIsometricOffAxis2Right = 32,
    /// <summary>
    /// 第二轴向上等轴测视图
    /// </summary>
    msoCameraIsometricOffAxis2Top = 33,
    /// <summary>
    /// 第三轴向左等轴测视图
    /// </summary>
    msoCameraIsometricOffAxis3Left = 34,
    /// <summary>
    /// 第三轴向右等轴测视图
    /// </summary>
    msoCameraIsometricOffAxis3Right = 35,
    /// <summary>
    /// 第三轴向下等轴测视图
    /// </summary>
    msoCameraIsometricOffAxis3Bottom = 36,
    /// <summary>
    /// 第四轴向左等轴测视图
    /// </summary>
    msoCameraIsometricOffAxis4Left = 37,
    /// <summary>
    /// 第四轴向右等轴测视图
    /// </summary>
    msoCameraIsometricOffAxis4Right = 38,
    /// <summary>
    /// 第四轴向下等轴测视图
    /// </summary>
    msoCameraIsometricOffAxis4Bottom = 39,
    /// <summary>
    /// 左上角斜角视图
    /// </summary>
    msoCameraObliqueTopLeft = 40,
    /// <summary>
    /// 顶部斜角视图
    /// </summary>
    msoCameraObliqueTop = 41,
    /// <summary>
    /// 右上角斜角视图
    /// </summary>
    msoCameraObliqueTopRight = 42,
    /// <summary>
    /// 左侧斜角视图
    /// </summary>
    msoCameraObliqueLeft = 43,
    /// <summary>
    /// 右侧斜角视图
    /// </summary>
    msoCameraObliqueRight = 44,
    /// <summary>
    /// 左下角斜角视图
    /// </summary>
    msoCameraObliqueBottomLeft = 45,
    /// <summary>
    /// 底部斜角视图
    /// </summary>
    msoCameraObliqueBottom = 46,
    /// <summary>
    /// 右下角斜角视图
    /// </summary>
    msoCameraObliqueBottomRight = 47,
    /// <summary>
    /// 前面透视视图
    /// </summary>
    msoCameraPerspectiveFront = 48,
    /// <summary>
    /// 左侧透视视图
    /// </summary>
    msoCameraPerspectiveLeft = 49,
    /// <summary>
    /// 右侧透视视图
    /// </summary>
    msoCameraPerspectiveRight = 50,
    /// <summary>
    /// 顶部透视视图
    /// </summary>
    msoCameraPerspectiveAbove = 51,
    /// <summary>
    /// 底部透视视图
    /// </summary>
    msoCameraPerspectiveBelow = 52,
    /// <summary>
    /// 面向左上方的透视视图
    /// </summary>
    msoCameraPerspectiveAboveLeftFacing = 53,
    /// <summary>
    /// 面向右上方的透视视图
    /// </summary>
    msoCameraPerspectiveAboveRightFacing = 54,
    /// <summary>
    /// 对比左侧透视视图
    /// </summary>
    msoCameraPerspectiveContrastingLeftFacing = 55,
    /// <summary>
    /// 对比右侧透视视图
    /// </summary>
    msoCameraPerspectiveContrastingRightFacing = 56,
    /// <summary>
    /// 英雄式左侧透视视图
    /// </summary>
    msoCameraPerspectiveHeroicLeftFacing = 57,
    /// <summary>
    /// 英雄式右侧透视视图
    /// </summary>
    msoCameraPerspectiveHeroicRightFacing = 58,
    /// <summary>
    /// 极端英雄式左侧透视视图
    /// </summary>
    msoCameraPerspectiveHeroicExtremeLeftFacing = 59,
    /// <summary>
    /// 极端英雄式右侧透视视图
    /// </summary>
    msoCameraPerspectiveHeroicExtremeRightFacing = 60,
    /// <summary>
    /// 放松透视视图
    /// </summary>
    msoCameraPerspectiveRelaxed = 61,
    /// <summary>
    /// 中度放松透视视图
    /// </summary>
    msoCameraPerspectiveRelaxedModerately = 62
}
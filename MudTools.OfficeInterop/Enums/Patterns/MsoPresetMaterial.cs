namespace MudTools.OfficeInterop;

/// <summary>
/// 指定在Office应用程序中应用于三维形状的预设材料类型
/// </summary>
public enum MsoPresetMaterial
{
    /// <summary>
    /// 混合材料类型
    /// </summary>
    msoPresetMaterialMixed = -2,
    
    /// <summary>
    /// 哑光材料
    /// </summary>
    msoMaterialMatte = 1,
    
    /// <summary>
    /// 塑料材料
    /// </summary>
    msoMaterialPlastic = 2,
    
    /// <summary>
    /// 金属材料
    /// </summary>
    msoMaterialMetal = 3,
    
    /// <summary>
    /// 线框材料
    /// </summary>
    msoMaterialWireFrame = 4,
    
    /// <summary>
    /// 哑光材料2
    /// </summary>
    msoMaterialMatte2 = 5,
    
    /// <summary>
    /// 塑料材料2
    /// </summary>
    msoMaterialPlastic2 = 6,
    
    /// <summary>
    /// 金属材料2
    /// </summary>
    msoMaterialMetal2 = 7,
    
    /// <summary>
    /// 暖哑光材料
    /// </summary>
    msoMaterialWarmMatte = 8,
    
    /// <summary>
    /// 半透明粉末材料
    /// </summary>
    msoMaterialTranslucentPowder = 9,
    
    /// <summary>
    /// 粉末材料
    /// </summary>
    msoMaterialPowder = 10,
    
    /// <summary>
    /// 深色边缘材料
    /// </summary>
    msoMaterialDarkEdge = 11,
    
    /// <summary>
    /// 软边缘材料
    /// </summary>
    msoMaterialSoftEdge = 12,
    
    /// <summary>
    /// 透明材料
    /// </summary>
    msoMaterialClear = 13,
    
    /// <summary>
    /// 平面材料
    /// </summary>
    msoMaterialFlat = 14,
    
    /// <summary>
    /// 软金属材料
    /// </summary>
    msoMaterialSoftMetal = 15
}
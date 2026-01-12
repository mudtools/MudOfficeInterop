namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定对象水平溢出处理方式的枚举
/// 当对象内容超出容器水平边界时，指定如何处理溢出部分
/// </summary>
public enum XlOartHorizontalOverflow
{
    /// <summary>
    /// 溢出显示
    /// 对象内容超出容器水平边界时仍然显示，不进行裁剪
    /// </summary>
    xlOartHorizontalOverflowOverflow,
    
    /// <summary>
    /// 裁剪隐藏
    /// 对象内容超出容器水平边界时进行裁剪，只显示容器范围内的内容
    /// </summary>
    xlOartHorizontalOverflowClip
}
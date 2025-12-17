//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定要返回切片的哪个位置坐标
/// </summary>
public enum XlPieSliceIndex
{
    /// <summary>
    /// 切片圆周上最外侧逆时针方向的点
    /// </summary>
    xlOuterCounterClockwisePoint = 1,

    /// <summary>
    /// 切片圆周上外侧中心点
    /// </summary>
    xlOuterCenterPoint,

    /// <summary>
    /// 切片圆周上最外侧顺时针方向的点
    /// </summary>
    xlOuterClockwisePoint,

    /// <summary>
    /// 切片最顺时针半径的中点
    /// </summary>
    xlMidClockwiseRadiusPoint,

    /// <summary>
    /// 饼图切片的中心点
    /// </summary>
    xlCenterPoint,

    /// <summary>
    /// 切片最逆时针半径的中点
    /// </summary>
    xlMidCounterClockwiseRadiusPoint,

    /// <summary>
    /// 环形图切片最顺时针半径的最内侧点
    /// </summary>
    xlInnerClockwisePoint,

    /// <summary>
    /// 环形图切片的最内侧中心点
    /// </summary>
    xlInnerCenterPoint,

    /// <summary>
    /// 环形图切片最逆时针半径的最内侧点
    /// </summary>
    xlInnerCounterClockwisePoint
}
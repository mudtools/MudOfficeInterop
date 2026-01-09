//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 图标集枚举
/// 用于指定条件格式中使用的图标集类型
/// </summary>
public enum XlIconSet
{
    /// <summary>
    /// 自定义图标集
    /// 用户自定义的图标集
    /// </summary>
    xlCustomSet = -1,

    /// <summary>
    /// 3个箭头图标集
    /// 包含绿色向上箭头、黄色横向箭头和红色向下箭头
    /// </summary>
    xl3Arrows = 1,

    /// <summary>
    /// 3个灰色箭头图标集
    /// 包含灰色向上箭头、灰色横向箭头和灰色向下箭头
    /// </summary>
    xl3ArrowsGray = 2,

    /// <summary>
    /// 3个旗帜图标集
    /// 包含绿色、黄色和红色旗帜
    /// </summary>
    xl3Flags = 3,

    /// <summary>
    /// 3个交通灯图标集（样式1）
    /// 包含绿色圆形、黄色三角形和红色圆形
    /// </summary>
    xl3TrafficLights1 = 4,

    /// <summary>
    /// 3个交通灯图标集（样式2）
    /// 包含绿色、黄色和红色交通灯
    /// </summary>
    xl3TrafficLights2 = 5,

    /// <summary>
    /// 3个标志图标集
    /// 包含绿色对号、黄色感叹号和红色叉号
    /// </summary>
    xl3Signs = 6,

    /// <summary>
    /// 3个符号图标集
    /// 包含笑脸、中性脸和哭脸
    /// </summary>
    xl3Symbols = 7,

    /// <summary>
    /// 3个符号图标集（样式2）
    /// 包含笑脸、中性脸和哭脸（略有不同）
    /// </summary>
    xl3Symbols2 = 8,

    /// <summary>
    /// 4个箭头图标集
    /// 包含绿色向上箭头、黄色向上箭头、黄色向下箭头和红色向下箭头
    /// </summary>
    xl4Arrows = 9,

    /// <summary>
    /// 4个灰色箭头图标集
    /// 包含不同方向的灰色箭头
    /// </summary>
    xl4ArrowsGray = 10,

    /// <summary>
    /// 4个红到黑图标集
    /// 包含红色、浅红色、灰色和黑色圆圈
    /// </summary>
    xl4RedToBlack = 11,

    /// <summary>
    /// 4个CRV图标集
    /// 包含柱状图样式的图标
    /// </summary>
    xl4CRV = 12,

    /// <summary>
    /// 4个交通灯图标集
    /// 包含不同颜色的交通灯
    /// </summary>
    xl4TrafficLights = 13,

    /// <summary>
    /// 5个箭头图标集
    /// 包含5个不同方向的箭头
    /// </summary>
    xl5Arrows = 14,

    /// <summary>
    /// 5个灰色箭头图标集
    /// 包含5个不同方向的灰色箭头
    /// </summary>
    xl5ArrowsGray = 15,

    /// <summary>
    /// 5个CRV图标集
    /// 包含5个柱状图样式的图标
    /// </summary>
    xl5CRV = 16,

    /// <summary>
    /// 5个四分之一图标集
    /// 包含表示0%、25%、50%、75%和100%的图标
    /// </summary>
    xl5Quarters = 17,

    /// <summary>
    /// 3个星形图标集
    /// 包含不同数量的星形图标
    /// </summary>
    xl3Stars = 18,

    /// <summary>
    /// 3个三角形图标集
    /// 包含绿色向上三角形、黄色横向三角形和红色向下三角形
    /// </summary>
    xl3Triangles = 19,

    /// <summary>
    /// 5个方框图标集
    /// 包含5个不同填充程度的方框
    /// </summary>
    xl5Boxes = 20
}
namespace MudTools.OfficeInterop;

/// &lt;summary&gt;
/// 指定形状线条端点的箭头样式
/// &lt;/summary&gt;
public enum MsoArrowheadStyle
{
    /// &lt;summary&gt;
    /// 仅用于持久化，表示混合样式
    /// &lt;/summary&gt;
    msoArrowheadStyleMixed = -2,

    /// &lt;summary&gt;
    /// 无箭头
    /// &lt;/summary&gt;
    msoArrowheadNone = 1,

    /// &lt;summary&gt;
    /// 三角形箭头
    /// &lt;/summary&gt;
    msoArrowheadTriangle = 2,

    /// &lt;summary&gt;
    /// 开放式箭头
    /// &lt;/summary&gt;
    msoArrowheadOpen = 3,

    /// &lt;summary&gt;
    /// 隐形箭头（尖锐的三角形变体）
    /// &lt;/summary&gt;
    msoArrowheadStealth = 4,

    /// &lt;summary&gt;
    /// 菱形箭头
    /// &lt;/summary&gt;
    msoArrowheadDiamond = 5,

    /// &lt;summary&gt;
    /// 椭圆形箭头
    /// &lt;/summary&gt;
    msoArrowheadOval = 6
}
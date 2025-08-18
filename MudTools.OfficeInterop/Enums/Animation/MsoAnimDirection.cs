//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定动画方向的枚举，用于PowerPoint动画效果的方向设置
/// </summary>
public enum MsoAnimDirection
{
    /// <summary>
    /// 无方向
    /// </summary>
    msoAnimDirectionNone = 0,

    /// <summary>
    /// 向上方向
    /// </summary>
    msoAnimDirectionUp = 1,

    /// <summary>
    /// 向右方向
    /// </summary>
    msoAnimDirectionRight = 2,

    /// <summary>
    /// 向下方向
    /// </summary>
    msoAnimDirectionDown = 3,

    /// <summary>
    /// 向左方向
    /// </summary>
    msoAnimDirectionLeft = 4,

    /// <summary>
    /// 左上角方向
    /// </summary>
    msoAnimDirectionTopLeft = 5,

    /// <summary>
    /// 右上角方向
    /// </summary>
    msoAnimDirectionTopRight = 6,

    /// <summary>
    /// 右下角方向
    /// </summary>
    msoAnimDirectionBottomRight = 7,

    /// <summary>
    /// 左下角方向
    /// </summary>
    msoAnimDirectionBottomLeft = 8,

    /// <summary>
    /// 水平方向
    /// </summary>
    msoAnimDirectionHorizontal = 9,

    /// <summary>
    /// 垂直方向
    /// </summary>
    msoAnimDirectionVertical = 10,

    /// <summary>
    /// 顺时针方向
    /// </summary>
    msoAnimDirectionClockwise = 11,

    /// <summary>
    /// 逆时针方向
    /// </summary>
    msoAnimDirectionCounterclockwise = 12,

    /// <summary>
    /// 水平向内
    /// </summary>
    msoAnimDirectionHorizontalIn = 13,

    /// <summary>
    /// 水平向外
    /// </summary>
    msoAnimDirectionHorizontalOut = 14,

    /// <summary>
    /// 垂直向内
    /// </summary>
    msoAnimDirectionVerticalIn = 15,

    /// <summary>
    /// 垂直向外
    /// </summary>
    msoAnimDirectionVerticalOut = 16,

    /// <summary>
    /// 轻微动画
    /// </summary>
    msoAnimDirectionSlightly = 17,

    /// <summary>
    /// 中心方向
    /// </summary>
    msoAnimDirectionCenter = 18,

    /// <summary>
    /// 轻微向内
    /// </summary>
    msoAnimDirectionInSlightly = 19,

    /// <summary>
    /// 轻微向外
    /// </summary>
    msoAnimDirectionOutSlightly = 20,

    /// <summary>
    /// 中心向内
    /// </summary>
    msoAnimDirectionInCenter = 21,

    /// <summary>
    /// 中心向外
    /// </summary>
    msoAnimDirectionOutCenter = 22,

    /// <summary>
    /// 底部向内
    /// </summary>
    msoAnimDirectionInBottom = 23,

    /// <summary>
    /// 底部向外
    /// </summary>
    msoAnimDirectionOutBottom = 24,

    /// <summary>
    /// 跨越动画
    /// </summary>
    msoAnimDirectionAcross = 25,

    /// <summary>
    /// 底部方向
    /// </summary>
    msoAnimDirectionBottom = 26,

    /// <summary>
    /// 顶部方向
    /// </summary>
    msoAnimDirectionTop = 27,

    /// <summary>
    /// 从左侧开始
    /// </summary>
    msoAnimDirectionFromLeft = 28,

    /// <summary>
    /// 从右侧开始
    /// </summary>
    msoAnimDirectionFromRight = 29,

    /// <summary>
    /// 从顶部开始
    /// </summary>
    msoAnimDirectionFromTop = 30,

    /// <summary>
    /// 从底部开始
    /// </summary>
    msoAnimDirectionFromBottom = 31,

    /// <summary>
    /// 从左上角开始
    /// </summary>
    msoAnimDirectionFromTopLeft = 32,

    /// <summary>
    /// 从右上角开始
    /// </summary>
    msoAnimDirectionFromTopRight = 33,

    /// <summary>
    /// 从左下角开始
    /// </summary>
    msoAnimDirectionFromBottomLeft = 34,

    /// <summary>
    /// 从右下角开始
    /// </summary>
    msoAnimDirectionFromBottomRight = 35,

    /// <summary>
    /// 向右结束
    /// </summary>
    msoAnimDirectionToRight = 36,

    /// <summary>
    /// 向左结束
    /// </summary>
    msoAnimDirectionToLeft = 37,

    /// <summary>
    /// 向上结束
    /// </summary>
    msoAnimDirectionToTop = 38,

    /// <summary>
    /// 向下结束
    /// </summary>
    msoAnimDirectionToBottom = 39,

    /// <summary>
    /// 向左上角结束
    /// </summary>
    msoAnimDirectionToTopLeft = 40,

    /// <summary>
    /// 向右上角结束
    /// </summary>
    msoAnimDirectionToTopRight = 41,

    /// <summary>
    /// 向左下角结束
    /// </summary>
    msoAnimDirectionToBottomLeft = 42,

    /// <summary>
    /// 向右下角结束
    /// </summary>
    msoAnimDirectionToBottomRight = 43
}
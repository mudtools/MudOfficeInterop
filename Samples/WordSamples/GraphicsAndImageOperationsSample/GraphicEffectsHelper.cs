//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace GraphicsAndImageOperationsSample
{
    /// <summary>
    /// 图形效果助手类
    /// </summary>
    public class GraphicEffectsHelper
    {
        /// <summary>
        /// 阴影效果类型枚举
        /// </summary>
        public enum ShadowEffectType
        {
            /// <summary>
            /// 外阴影
            /// </summary>
            OuterShadow,

            /// <summary>
            /// 内阴影
            /// </summary>
            InnerShadow,

            /// <summary>
            /// 透视外阴影
            /// </summary>
            PerspectiveOuterShadow,

            /// <summary>
            /// 透视内阴影
            /// </summary>
            PerspectiveInnerShadow
        }

        /// <summary>
        /// 发光效果类型枚举
        /// </summary>
        public enum GlowEffectType
        {
            /// <summary>
            /// 小发光
            /// </summary>
            Small,

            /// <summary>
            /// 中等发光
            /// </summary>
            Medium,

            /// <summary>
            /// 大发光
            /// </summary>
            Large,

            /// <summary>
            /// 自定义发光
            /// </summary>
            Custom
        }

        /// <summary>
        /// 三维效果类型枚举
        /// </summary>
        public enum ThreeDEffectType
        {
            /// <summary>
            /// 简单三维
            /// </summary>
            Simple,

            /// <summary>
            /// 复杂三维
            /// </summary>
            Complex,

            /// <summary>
            /// 金属质感
            /// </summary>
            Metallic,

            /// <summary>
            /// 塑料质感
            /// </summary>
            Plastic
        }

        /// <summary>
        /// 为图形添加阴影效果
        /// </summary>
        /// <param name="shape">图形对象</param>
        /// <param name="effectType">阴影效果类型</param>
        /// <param name="color">阴影颜色</param>
        /// <param name="offsetX">X轴偏移</param>
        /// <param name="offsetY">Y轴偏移</param>
        /// <param name="blur">模糊度</param>
        /// <param name="transparency">透明度</param>
        /// <param name="size">阴影大小</param>
        public void ApplyShadowEffect(
            IWordShape shape,
            ShadowEffectType effectType,
            WdColor color,
            float offsetX = 3,
            float offsetY = 3,
            float blur = 5,
            float transparency = 0,
            float size = 100)
        {
            if (shape?.Shadow == null) return;

            try
            {
                // 设置阴影可见性
                shape.Shadow.Visible = true;

                // 根据效果类型设置阴影样式
                switch (effectType)
                {
                    case ShadowEffectType.OuterShadow:
                        shape.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
                        break;
                    case ShadowEffectType.InnerShadow:
                        shape.Shadow.Style = MsoShadowStyle.msoShadowStyleInnerShadow;
                        break;
                    case ShadowEffectType.PerspectiveOuterShadow:
                        shape.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
                        break;
                    case ShadowEffectType.PerspectiveInnerShadow:
                        shape.Shadow.Style = MsoShadowStyle.msoShadowStyleInnerShadow;
                        break;
                }

                // 设置阴影属性
                shape.Shadow.ForeColor.RGB = (int)color;
                shape.Shadow.OffsetX = offsetX;
                shape.Shadow.OffsetY = offsetY;
                shape.Shadow.Blur = blur;
                shape.Shadow.Transparency = transparency;
                shape.Shadow.Size = size;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用阴影效果时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 为图形添加发光效果
        /// </summary>
        /// <param name="shape">图形对象</param>
        /// <param name="effectType">发光效果类型</param>
        /// <param name="color">发光颜色</param>
        /// <param name="radius">发光半径</param>
        /// <param name="transparency">透明度</param>
        public void ApplyGlowEffect(
            IWordShape shape,
            GlowEffectType effectType,
            WdColor color,
            float radius = 5,
            float transparency = 0)
        {
            if (shape?.Glow == null) return;

            try
            {
                // 根据效果类型设置发光半径
                switch (effectType)
                {
                    case GlowEffectType.Small:
                        radius = 3;
                        break;
                    case GlowEffectType.Medium:
                        radius = 7;
                        break;
                    case GlowEffectType.Large:
                        radius = 12;
                        break;
                }

                // 设置发光属性
                shape.Glow.Radius = radius;
                shape.Glow.Color.RGB = (int)color;
                shape.Glow.Transparency = transparency;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用发光效果时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 为图形添加柔化边缘效果
        /// </summary>
        /// <param name="shape">图形对象</param>
        /// <param name="radius">柔化半径</param>
        public void ApplySoftEdgeEffect(IWordShape shape, float radius)
        {
            if (shape?.SoftEdge == null) return;

            try
            {
                if (radius <= 0)
                {
                    shape.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeTypeNone;
                }
                else if (radius <= 2)
                {
                    shape.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeType1;
                }
                else if (radius <= 5)
                {
                    shape.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeType2;
                }
                else if (radius <= 10)
                {
                    shape.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeType3;
                }
                else
                {
                    shape.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeType6;
                }

                shape.SoftEdge.Radius = radius;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用柔化边缘效果时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 为图形添加反射效果
        /// </summary>
        /// <param name="shape">图形对象</param>
        /// <param name="transparency">透明度</param>
        /// <param name="size">反射大小</param>
        /// <param name="offset">反射偏移</param>
        /// <param name="blur">模糊度</param>
        public void ApplyReflectionEffect(
            IWordShape shape,
            float transparency = 0.5f,
            float size = 50,
            float offset = 0,
            float blur = 5)
        {
            if (shape?.Reflection == null) return;

            try
            {
                shape.Reflection.Type = MsoReflectionType.msoReflectionType1; // 默认反射类型
                shape.Reflection.Transparency = transparency;
                shape.Reflection.Size = size;
                shape.Reflection.Offset = offset;
                shape.Reflection.Blur = blur;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用反射效果时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 为图形添加三维旋转效果
        /// </summary>
        /// <param name="shape">图形对象</param>
        /// <param name="rotationX">X轴旋转角度</param>
        /// <param name="rotationY">Y轴旋转角度</param>
        /// <param name="rotationZ">Z轴旋转角度</param>
        /// <param name="perspective">透视效果</param>
        public void Apply3DRotationEffect(
            IWordShape shape,
            float rotationX = 0,
            float rotationY = 0,
            float rotationZ = 0,
            float perspective = 0)
        {
            if (shape?.ThreeD == null) return;

            try
            {
                shape.ThreeD.Visible = true;
                shape.ThreeD.RotationX = rotationX;
                shape.ThreeD.RotationY = rotationY;
                shape.ThreeD.RotationZ = rotationZ;

                if (perspective > 0)
                {
                    shape.ThreeD.Perspective = true;
                    shape.ThreeD.Perspective = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用三维旋转效果时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 为图形添加三维格式效果
        /// </summary>
        /// <param name="shape">图形对象</param>
        /// <param name="effectType">三维效果类型</param>
        /// <param name="depth">深度</param>
        /// <param name="extrusionColor">挤出颜色</param>
        /// <param name="contourColor">轮廓颜色</param>
        /// <param name="contourWidth">轮廓宽度</param>
        public void Apply3DFormatEffect(
            IWordShape shape,
            ThreeDEffectType effectType,
            float depth = 2,
            WdColor extrusionColor = WdColor.wdColorAutomatic,
            WdColor contourColor = WdColor.wdColorAutomatic,
            float contourWidth = 1)
        {
            if (shape?.ThreeD == null) return;

            try
            {
                shape.ThreeD.Visible = true;
                shape.ThreeD.Depth = depth;

                // 设置挤出颜色
                if (extrusionColor != WdColor.wdColorAutomatic)
                {
                    shape.ThreeD.ExtrusionColor.RGB = (int)extrusionColor;
                }

                // 设置轮廓
                if (contourColor != WdColor.wdColorAutomatic)
                {
                    shape.ThreeD.ContourColor.RGB = (int)contourColor;
                }
                shape.ThreeD.ContourWidth = contourWidth;

                // 根据效果类型设置其他属性
                switch (effectType)
                {
                    case ThreeDEffectType.Simple:
                        shape.ThreeD.BevelTopType = MsoBevelType.msoBevelCircle;
                        shape.ThreeD.BevelTopInset = 3;
                        shape.ThreeD.BevelTopDepth = 2;
                        break;
                    case ThreeDEffectType.Complex:
                        shape.ThreeD.BevelTopType = MsoBevelType.msoBevelArtDeco;
                        shape.ThreeD.BevelTopInset = 5;
                        shape.ThreeD.BevelTopDepth = 4;
                        break;
                    case ThreeDEffectType.Metallic:
                        shape.ThreeD.BevelTopType = MsoBevelType.msoBevelCoolSlant;
                        shape.ThreeD.BevelTopInset = 4;
                        shape.ThreeD.BevelTopDepth = 3;
                        break;
                    case ThreeDEffectType.Plastic:
                        shape.ThreeD.BevelTopType = MsoBevelType.msoBevelDivot;
                        shape.ThreeD.BevelTopInset = 2;
                        shape.ThreeD.BevelTopDepth = 1;
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用三维格式效果时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 应用预设效果组合
        /// </summary>
        /// <param name="shape">图形对象</param>
        /// <param name="presetName">预设名称</param>
        public void ApplyPresetEffect(IWordShape shape, string presetName)
        {
            try
            {
                switch (presetName.ToLower())
                {
                    case "glow":
                        ApplyGlowEffect(shape, GlowEffectType.Medium, WdColor.wdColorBlue);
                        break;
                    case "shadow":
                        ApplyShadowEffect(shape, ShadowEffectType.OuterShadow, WdColor.wdColorGray50);
                        break;
                    case "3d":
                        Apply3DFormatEffect(shape, ThreeDEffectType.Simple);
                        break;
                    case "softedge":
                        ApplySoftEdgeEffect(shape, 5);
                        break;
                    case "reflection":
                        ApplyReflectionEffect(shape);
                        break;
                    case "professional":
                        ApplyShadowEffect(shape, ShadowEffectType.OuterShadow, WdColor.wdColorGray50, 2, 2, 3);
                        ApplyGlowEffect(shape, GlowEffectType.Small, WdColor.wdColorBlue, 2);
                        ApplySoftEdgeEffect(shape, 2);
                        break;
                    case "dramatic":
                        ApplyShadowEffect(shape, ShadowEffectType.PerspectiveOuterShadow, WdColor.wdColorBlack, 5, 5, 10);
                        Apply3DFormatEffect(shape, ThreeDEffectType.Complex, 5);
                        ApplyGlowEffect(shape, GlowEffectType.Large, WdColor.wdColorGold, 10);
                        break;
                    default:
                        Console.WriteLine($"未知的预设效果: {presetName}");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用预设效果时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 移除所有效果
        /// </summary>
        /// <param name="shape">图形对象</param>
        public void RemoveAllEffects(IWordShape shape)
        {
            try
            {
                // 移除阴影效果
                if (shape?.Shadow != null)
                {
                    shape.Shadow.Visible = false;
                }

                // 移除发光效果
                if (shape?.Glow != null)
                {
                    shape.Glow.Radius = 0;
                }

                // 移除柔化边缘效果
                if (shape?.SoftEdge != null)
                {
                    shape.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeTypeNone;
                }

                // 移除反射效果
                if (shape?.Reflection != null)
                {
                    shape.Reflection.Type = MsoReflectionType.msoReflectionTypeNone;
                }

                // 移除三维效果
                if (shape?.ThreeD != null)
                {
                    shape.ThreeD.Visible = false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"移除效果时出错: {ex.Message}");
            }
        }
    }
}
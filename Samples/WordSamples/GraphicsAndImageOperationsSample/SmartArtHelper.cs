//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;

namespace GraphicsAndImageOperationsSample
{
    /// <summary>
    /// SmartArt图形助手类
    /// </summary>
    public class SmartArtHelper
    {
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public SmartArtHelper(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 创建列表类型的SmartArt
        /// </summary>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <returns>SmartArt图形对象</returns>
        public IWordShape CreateListSmartArt(float left, float top, float width, float height)
        {
            var range = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
            var smartArtShape = _document.Shapes.AddSmartArt(
                null, // 需要传入具体的SmartArtLayout对象，暂时传null
                left, top, width, height);
            return smartArtShape;
        }

        /// <summary>
        /// 创建循环类型的SmartArt
        /// </summary>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <returns>SmartArt图形对象</returns>
        public IWordShape CreateCycleSmartArt(float left, float top, float width, float height)
        {
            var range = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
            var smartArtShape = _document.Shapes.AddSmartArt(
                null, // 需要传入具体的SmartArtLayout对象，暂时传null
                left, top, width, height);
            return smartArtShape;
        }

        /// <summary>
        /// 创建层次结构类型的SmartArt
        /// </summary>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <returns>SmartArt图形对象</returns>
        public IWordShape CreateHierarchySmartArt(float left, float top, float width, float height)
        {
            var range = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
            var smartArtShape = _document.Shapes.AddSmartArt(
                null, // 需要传入具体的SmartArtLayout对象，暂时传null
                left, top, width, height);
            return smartArtShape;
        }

        /// <summary>
        /// 创建流程图类型的SmartArt
        /// </summary>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <returns>SmartArt图形对象</returns>
        public IWordShape CreateProcessSmartArt(float left, float top, float width, float height)
        {
            var range = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
            var smartArtShape = _document.Shapes.AddSmartArt(
                null, // 需要传入具体的SmartArtLayout对象，暂时传null
                left, top, width, height);
            return smartArtShape;
        }

        /// <summary>
        /// 设置SmartArt节点文本
        /// </summary>
        /// <param name="smartArtShape">SmartArt图形对象</param>
        /// <param name="nodeTexts">节点文本列表</param>
        public void SetNodeTexts(IWordShape smartArtShape, List<string> nodeTexts)
        {
            if (smartArtShape?.SmartArt?.AllNodes != null && nodeTexts != null)
            {
                var smartArt = smartArtShape.SmartArt;
                for (int i = 0; i < Math.Min(nodeTexts.Count, smartArt.AllNodes.Count); i++)
                {
                    try
                    {
                        smartArt.AllNodes[i + 1].TextFrame.TextRange.Text = nodeTexts[i];
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"设置节点 {i + 1} 文本时出错: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// 添加SmartArt节点
        /// </summary>
        /// <param name="smartArtShape">SmartArt图形对象</param>
        /// <param name="text">节点文本</param>
        /// <returns>添加的节点</returns>
        public object AddNode(IWordShape smartArtShape, string text)
        {
            if (smartArtShape?.SmartArt != null)
            {
                try
                {
                    var smartArt = smartArtShape.SmartArt;
                    // 添加节点的逻辑
                    // 注意：实际实现需要根据具体API进行调整
                    Console.WriteLine("节点添加功能演示");
                    return new object();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"添加节点时出错: {ex.Message}");
                    return null;
                }
            }
            return null;
        }

        /// <summary>
        /// 删除SmartArt节点
        /// </summary>
        /// <param name="smartArtShape">SmartArt图形对象</param>
        /// <param name="nodeIndex">节点索引</param>
        public void RemoveNode(IWordShape smartArtShape, int nodeIndex)
        {
            if (smartArtShape?.SmartArt?.AllNodes != null)
            {
                try
                {
                    // 删除节点的逻辑
                    // 注意：实际实现需要根据具体API进行调整
                    Console.WriteLine($"删除节点 {nodeIndex} 功能演示");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"删除节点时出错: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 设置SmartArt颜色样式
        /// </summary>
        /// <param name="smartArtShape">SmartArt图形对象</param>
        /// <param name="colorIndex">颜色索引</param>
        public void SetColorStyle(IWordShape smartArtShape, int colorIndex)
        {
            if (smartArtShape?.SmartArt != null && _document?.Application?.SmartArtColors != null)
            {
                try
                {
                    var smartArt = smartArtShape.SmartArt;
                    if (_document.Application.SmartArtColors.Count >= colorIndex)
                    {
                        smartArt.Color = _document.Application.SmartArtColors[colorIndex];
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"设置颜色样式时出错: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 设置SmartArt布局样式
        /// </summary>
        /// <param name="smartArtShape">SmartArt图形对象</param>
        /// <param name="layoutIndex">布局索引</param>
        public void SetLayoutStyle(IWordShape smartArtShape, int layoutIndex)
        {
            if (smartArtShape?.SmartArt != null && _document?.Application?.SmartArtLayouts != null)
            {
                try
                {
                    var smartArt = smartArtShape.SmartArt;
                    if (_document.Application.SmartArtLayouts.Count >= layoutIndex)
                    {
                        smartArt.Layout = _document.Application.SmartArtLayouts[layoutIndex];
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"设置布局样式时出错: {ex.Message}");
                }
            }
        }
    }
}
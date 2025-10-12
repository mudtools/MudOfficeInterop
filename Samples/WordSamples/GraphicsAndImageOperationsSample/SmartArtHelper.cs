using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                MsoSmartArtDefaultConstants.msoSmartArtDefaultList,
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
                MsoSmartArtDefaultConstants.msoSmartArtDefaultCycle,
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
                MsoSmartArtDefaultConstants.msoSmartArtDefaultHierarchy,
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
                MsoSmartArtDefaultConstants.msoSmartArtDefaultProcess,
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
            if (smartArtShape?.SmartArt?.Parent?.SmartArtColors != null)
            {
                try
                {
                    var smartArt = smartArtShape.SmartArt;
                    if (smartArt.Parent.SmartArtColors.Count >= colorIndex)
                    {
                        smartArt.Color = smartArt.Parent.SmartArtColors[colorIndex];
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
            if (smartArtShape?.SmartArt?.Parent?.SmartArtLayouts != null)
            {
                try
                {
                    var smartArt = smartArtShape.SmartArt;
                    if (smartArt.Parent.SmartArtLayouts.Count >= layoutIndex)
                    {
                        smartArt.Layout = smartArt.Parent.SmartArtLayouts[layoutIndex];
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"设置布局样式时出错: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 设置SmartArt样式
        /// </summary>
        /// <param name="smartArtShape">SmartArt图形对象</param>
        /// <param name="styleIndex">样式索引</param>
        public void SetStyle(IWordShape smartArtShape, int styleIndex)
        {
            if (smartArtShape?.SmartArt?.Parent?.SmartArtQuickStyles != null)
            {
                try
                {
                    var smartArt = smartArtShape.SmartArt;
                    if (smartArt.Parent.SmartArtQuickStyles.Count >= styleIndex)
                    {
                        smartArt.QuickStyle = smartArt.Parent.SmartArtQuickStyles[styleIndex];
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"设置样式时出错: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 创建自定义SmartArt图表
        /// </summary>
        /// <param name="type">SmartArt类型</param>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <param name="nodeTexts">节点文本</param>
        /// <returns>SmartArt图形对象</returns>
        public IWordShape CreateCustomSmartArt(
            MsoSmartArtDefaultConstants type,
            float left, float top, float width, float height,
            List<string> nodeTexts)
        {
            try
            {
                var smartArtShape = _document.Shapes.AddSmartArt(type, left, top, width, height);
                SetNodeTexts(smartArtShape, nodeTexts);
                return smartArtShape;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建自定义SmartArt时出错: {ex.Message}");
                return null;
            }
        }
    }
}
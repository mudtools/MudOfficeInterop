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
    /// 图片处理器类
    /// </summary>
    public class ImageProcessor
    {
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public ImageProcessor(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 调整图片亮度
        /// </summary>
        /// <param name="inlineShape">内嵌图形对象</param>
        /// <param name="brightness">亮度值 (-1.0 到 1.0)</param>
        public void AdjustBrightness(IWordInlineShape inlineShape, float brightness)
        {
            if (inlineShape != null && inlineShape.PictureFormat != null)
            {
                inlineShape.PictureFormat.Brightness = brightness;
            }
        }

        /// <summary>
        /// 调整图片对比度
        /// </summary>
        /// <param name="inlineShape">内嵌图形对象</param>
        /// <param name="contrast">对比度值 (-1.0 到 1.0)</param>
        public void AdjustContrast(IWordInlineShape inlineShape, float contrast)
        {
            if (inlineShape != null && inlineShape.PictureFormat != null)
            {
                inlineShape.PictureFormat.Contrast = contrast;
            }
        }

        /// <summary>
        /// 设置图片尺寸
        /// </summary>
        /// <param name="inlineShape">内嵌图形对象</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <param name="maintainAspectRatio">是否保持纵横比</param>
        public void ResizeImage(IWordInlineShape inlineShape, float width, float height, bool maintainAspectRatio = true)
        {
            if (inlineShape != null)
            {
                inlineShape.Width = width;
                inlineShape.Height = height;
                inlineShape.LockAspectRatio = maintainAspectRatio;
            }
        }

        /// <summary>
        /// 设置图片替代文本
        /// </summary>
        /// <param name="inlineShape">内嵌图形对象</param>
        /// <param name="alternativeText">替代文本</param>
        public void SetAlternativeText(IWordInlineShape inlineShape, string alternativeText)
        {
            if (inlineShape != null)
            {
                inlineShape.AlternativeText = alternativeText;
            }
        }

        /// <summary>
        /// 将内嵌图片转换为浮动图片
        /// </summary>
        /// <param name="inlineShape">内嵌图形对象</param>
        /// <returns>浮动图形对象</returns>
        public IWordShape ConvertToFloatImage(IWordInlineShape inlineShape)
        {
            if (inlineShape != null)
            {
                return inlineShape.ConvertToShape();
            }
            return null;
        }

        /// <summary>
        /// 设置图片环绕方式
        /// </summary>
        /// <param name="shape">浮动图形对象</param>
        /// <param name="wrapType">环绕类型</param>
        public void SetWrapFormat(IWordShape shape, WdWrapType wrapType)
        {
            if (shape != null && shape.WrapFormat != null)
            {
                shape.WrapFormat.Type = wrapType;
            }
        }

        /// <summary>
        /// 批量处理图片
        /// </summary>
        /// <param name="imagePaths">图片路径列表</param>
        /// <param name="targetWidth">目标宽度</param>
        /// <param name="targetHeight">目标高度</param>
        /// <returns>处理后的内嵌图形列表</returns>
        public List<IWordInlineShape> ProcessImages(List<string> imagePaths, float targetWidth, float targetHeight)
        {
            var processedImages = new List<IWordInlineShape>();

            foreach (var imagePath in imagePaths)
            {
                try
                {
                    if (System.IO.File.Exists(imagePath))
                    {
                        var range = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                        var inlineShape = range.InlineShapes.AddPicture(imagePath);
                        ResizeImage(inlineShape, targetWidth, targetHeight);
                        SetAlternativeText(inlineShape, System.IO.Path.GetFileNameWithoutExtension(imagePath));
                        processedImages.Add(inlineShape);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"处理图片 {imagePath} 时出错: {ex.Message}");
                }
            }

            return processedImages;
        }
    }
}
//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint 形状范围实现类
/// </summary>
internal class PowerPointShapeRange : IPowerPointShapeRange
{
    private readonly MsPowerPoint.ShapeRange _shapeRange;
    private bool _disposedValue;
    private IPowerPointTextFrame _textFrame;
    private IPowerPointFillFormat _fill;
    private IPowerPointLineFormat _line;
    private IPowerPointShadowFormat _shadow;
    private IPowerPointThreeDFormat _threeD;
    private IPowerPointTags _tags;

    /// <summary>
    /// 获取形状数量
    /// </summary>
    public int Count => _shapeRange.Count;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _shapeRange.Parent;

    /// <summary>
    /// 获取或设置形状范围名称
    /// </summary>
    public string Name
    {
        get => _shapeRange.Name;
        set => _shapeRange.Name = value;
    }

    /// <summary>
    /// 获取形状范围的左边缘位置
    /// </summary>
    public float Left
    {
        get => _shapeRange.Left;
        set => _shapeRange.Left = value;
    }

    /// <summary>
    /// 获取形状范围的上边缘位置
    /// </summary>
    public float Top
    {
        get => _shapeRange.Top;
        set => _shapeRange.Top = value;
    }

    /// <summary>
    /// 获取形状范围的宽度
    /// </summary>
    public float Width
    {
        get => _shapeRange.Width;
        set => _shapeRange.Width = value;
    }

    /// <summary>
    /// 获取形状范围的高度
    /// </summary>
    public float Height
    {
        get => _shapeRange.Height;
        set => _shapeRange.Height = value;
    }

    /// <summary>
    /// 获取形状范围的旋转角度
    /// </summary>
    public float Rotation
    {
        get => _shapeRange.Rotation;
        set => _shapeRange.Rotation = value;
    }

    /// <summary>
    /// 获取或设置形状范围的可见性
    /// </summary>
    public bool Visible
    {
        get => _shapeRange.Visible == Microsoft.Office.Core.MsoTriState.msoTrue;
        set => _shapeRange.Visible = value ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
    }

    /// <summary>
    /// 获取形状范围的锁定纵横比
    /// </summary>
    public bool LockAspectRatio
    {
        get => _shapeRange.LockAspectRatio == Microsoft.Office.Core.MsoTriState.msoTrue;
        set => _shapeRange.LockAspectRatio = value ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
    }

    /// <summary>
    /// 获取形状范围的Z轴顺序
    /// </summary>
    public int ZOrderPosition => _shapeRange.ZOrderPosition;

    /// <summary>
    /// 获取形状范围的文本框
    /// </summary>
    public IPowerPointTextFrame TextFrame
    {
        get
        {
            if (_textFrame == null)
            {
                _textFrame = new PowerPointTextFrame(_shapeRange.TextFrame);
            }
            return _textFrame;
        }
    }

    /// <summary>
    /// 获取形状范围的填充格式
    /// </summary>
    public IPowerPointFillFormat Fill
    {
        get
        {
            if (_fill == null)
            {
                _fill = new PowerPointFillFormat(_shapeRange.Fill);
            }
            return _fill;
        }
    }

    /// <summary>
    /// 获取形状范围的线条格式
    /// </summary>
    public IPowerPointLineFormat Line
    {
        get
        {
            if (_line == null)
            {
                _line = new PowerPointLineFormat(_shapeRange.Line);
            }
            return _line;
        }
    }

    /// <summary>
    /// 获取形状范围的阴影格式
    /// </summary>
    public IPowerPointShadowFormat Shadow
    {
        get
        {
            if (_shadow == null)
            {
                _shadow = new PowerPointShadowFormat(_shapeRange.Shadow);
            }
            return _shadow;
        }
    }

    /// <summary>
    /// 获取形状范围的三维格式
    /// </summary>
    public IPowerPointThreeDFormat ThreeD
    {
        get
        {
            if (_threeD == null)
            {
                _threeD = new PowerPointThreeDFormat(_shapeRange.ThreeD);
            }
            return _threeD;
        }
    }


    /// <summary>
    /// 获取形状范围的动画设置
    /// </summary>
    public IPowerPointAnimationSettings AnimationSettings
    {
        get
        {
            try
            {
                var animationSettings = _shapeRange.AnimationSettings;
                return animationSettings != null ? new PowerPointAnimationSettings(animationSettings) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取形状范围的标签集合
    /// </summary>
    public IPowerPointTags Tags
    {
        get
        {
            if (_tags == null)
            {
                _tags = new PowerPointTags(_shapeRange.Tags);
            }
            return _tags;
        }
    }

    /// <summary>
    /// 获取形状范围内的所有形状
    /// </summary>
    public IEnumerable<IPowerPointShape> Shapes
    {
        get
        {
            var shapes = new List<IPowerPointShape>();
            try
            {
                for (int i = 1; i <= _shapeRange.Count; i++)
                {
                    try
                    {
                        shapes.Add(new PowerPointShape(_shapeRange[i]));
                    }
                    catch
                    {
                        // 忽略获取单个形状失败的情况
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to enumerate shapes in shape range.", ex);
            }
            return shapes;
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="shapeRange">COM ShapeRange 对象</param>
    internal PowerPointShapeRange(MsPowerPoint.ShapeRange shapeRange)
    {
        _shapeRange = shapeRange ?? throw new ArgumentNullException(nameof(shapeRange));
        _disposedValue = false;
    }

    /// <summary>
    /// 选择形状范围
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    public void Select(bool replace = true)
    {
        try
        {
            _shapeRange.Select(replace ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to select shape range.", ex);
        }
    }

    /// <summary>
    /// 复制形状范围
    /// </summary>
    public void Copy()
    {
        try
        {
            _shapeRange.Copy();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to copy shape range.", ex);
        }
    }

    /// <summary>
    /// 剪切形状范围
    /// </summary>
    public void Cut()
    {
        try
        {
            _shapeRange.Cut();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to cut shape range.", ex);
        }
    }

    /// <summary>
    /// 删除形状范围
    /// </summary>
    public void Delete()
    {
        try
        {
            _shapeRange.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete shape range.", ex);
        }
    }


    /// <summary>
    /// 水平翻转形状范围
    /// </summary>
    public void FlipHorizontal()
    {
        try
        {
            _shapeRange.Flip(Microsoft.Office.Core.MsoFlipCmd.msoFlipHorizontal);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to flip shape range horizontally.", ex);
        }
    }

    /// <summary>
    /// 垂直翻转形状范围
    /// </summary>
    public void FlipVertical()
    {
        try
        {
            _shapeRange.Flip(Microsoft.Office.Core.MsoFlipCmd.msoFlipVertical);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to flip shape range vertically.", ex);
        }
    }

    /// <summary>
    /// 设置形状范围的Z轴顺序
    /// </summary>
    /// <param name="position">Z轴顺序位置</param>
    public void ZOrder(int position)
    {
        try
        {
            _shapeRange.ZOrder((Microsoft.Office.Core.MsoZOrderCmd)position);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set Z-order of shape range.", ex);
        }
    }

    /// <summary>
    /// 组合形状范围
    /// </summary>
    /// <returns>组合后的形状</returns>
    public IPowerPointShape Group()
    {
        try
        {
            var groupedShape = _shapeRange.Group();
            return new PowerPointShape(groupedShape);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to group shape range.", ex);
        }
    }

    /// <summary>
    /// 取消组合形状范围
    /// </summary>
    /// <returns>取消组合后的形状范围</returns>
    public IPowerPointShapeRange Ungroup()
    {
        try
        {
            var ungroupedRange = _shapeRange.Ungroup();
            return new PowerPointShapeRange(ungroupedRange);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to ungroup shape range.", ex);
        }
    }

    /// <summary>
    /// 对齐形状范围
    /// </summary>
    /// <param name="alignCmd">对齐命令</param>
    /// <param name="relativeToSlide">是否相对于幻灯片对齐</param>
    public void Align(int alignCmd, bool relativeToSlide = false)
    {
        try
        {
            _shapeRange.Align((Microsoft.Office.Core.MsoAlignCmd)alignCmd,
                relativeToSlide ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to align shape range.", ex);
        }
    }

    /// <summary>
    /// 分布形状范围
    /// </summary>
    /// <param name="distributeCmd">分布命令</param>
    /// <param name="relativeToSlide">是否相对于幻灯片分布</param>
    public void Distribute(int distributeCmd, bool relativeToSlide = false)
    {
        try
        {
            _shapeRange.Distribute((Microsoft.Office.Core.MsoDistributeCmd)distributeCmd,
                relativeToSlide ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to distribute shape range.", ex);
        }
    }

    /// <summary>
    /// 获取指定索引的形状
    /// </summary>
    /// <param name="index">形状索引</param>
    /// <returns>形状对象</returns>
    public IPowerPointShape Item(int index)
    {
        if (index < 1 || index > Count)
            throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

        try
        {
            var shape = _shapeRange[index];
            return new PowerPointShape(shape);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get shape at index {index}.", ex);
        }
    }

    /// <summary>
    /// 获取指定名称的形状
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <returns>形状对象</returns>
    public IPowerPointShape Item(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Shape name cannot be null or empty.", nameof(name));

        try
        {
            var shape = _shapeRange[name];
            return new PowerPointShape(shape);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get shape with name '{name}'.", ex);
        }
    }

    /// <summary>
    /// 根据条件查找形状
    /// </summary>
    /// <param name="predicate">查找条件</param>
    /// <returns>符合条件的形状列表</returns>
    public IEnumerable<IPowerPointShape> Find(Func<IPowerPointShape, bool> predicate)
    {
        if (predicate == null)
            throw new ArgumentNullException(nameof(predicate));

        var results = new List<IPowerPointShape>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var shape = Item(i);
                    if (predicate(shape))
                    {
                        results.Add(shape);
                    }
                }
                catch
                {
                    // 忽略获取或判断单个形状失败的情况
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to find shapes in shape range.", ex);
        }
        return results;
    }

    /// <summary>
    /// 获取形状范围的文本内容
    /// </summary>
    /// <returns>文本内容</returns>
    public string GetText()
    {
        try
        {
            return _shapeRange.TextFrame?.TextRange?.Text ?? string.Empty;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get text from shape range.", ex);
        }
    }

    /// <summary>
    /// 设置形状范围的文本内容
    /// </summary>
    /// <param name="text">文本内容</param>
    public void SetText(string text)
    {
        try
        {
            if (_shapeRange.TextFrame != null)
            {
                _shapeRange.TextFrame.TextRange.Text = text ?? string.Empty;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set text for shape range.", ex);
        }
    }

    /// <summary>
    /// 替换形状范围中的文本
    /// </summary>
    /// <param name="findText">查找文本</param>
    /// <param name="replaceText">替换文本</param>
    /// <returns>替换次数</returns>
    public int ReplaceText(string findText, string replaceText)
    {
        if (string.IsNullOrEmpty(findText))
            return 0;

        var count = 0;
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var shape = _shapeRange[i];
                    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        var replaceCount = shape.TextFrame.TextRange.Replace(findText, replaceText ?? string.Empty);
                        count += replaceCount.Count;
                    }
                }
                catch
                {
                    // 忽略单个形状替换失败的情况
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to replace text in shape range.", ex);
        }
        return count;
    }

    /// <summary>
    /// 添加文本到形状范围
    /// </summary>
    /// <param name="text">要添加的文本</param>
    public void AddText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return;

        try
        {
            var currentText = GetText();
            SetText(currentText + text);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add text to shape range.", ex);
        }
    }

    /// <summary>
    /// 清除形状范围的文本
    /// </summary>
    public void ClearText()
    {
        try
        {
            SetText(string.Empty);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to clear text from shape range.", ex);
        }
    }

    /// <summary>
    /// 设置形状范围的填充颜色
    /// </summary>
    /// <param name="color">颜色值</param>
    public void SetFillColor(int color)
    {
        try
        {
            _shapeRange.Fill.ForeColor.RGB = color;
            _shapeRange.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set fill color for shape range.", ex);
        }
    }

    /// <summary>
    /// 设置形状范围的线条颜色
    /// </summary>
    /// <param name="color">颜色值</param>
    public void SetLineColor(int color)
    {
        try
        {
            _shapeRange.Line.ForeColor.RGB = color;
            _shapeRange.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set line color for shape range.", ex);
        }
    }

    /// <summary>
    /// 设置形状范围的线条粗细
    /// </summary>
    /// <param name="weight">线条粗细</param>
    public void SetLineWeight(float weight)
    {
        try
        {
            _shapeRange.Line.Weight = weight;
            _shapeRange.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set line weight for shape range.", ex);
        }
    }

    /// <summary>
    /// 应用阴影效果
    /// </summary>
    /// <param name="shadowType">阴影类型</param>
    public void ApplyShadow(int shadowType = 1)
    {
        try
        {
            _shapeRange.Shadow.Visible = MsCore.MsoTriState.msoTrue;
            _shapeRange.Shadow.Type = (MsCore.MsoShadowType)shadowType;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply shadow to shape range.", ex);
        }
    }

    /// <summary>
    /// 应用三维效果
    /// </summary>
    /// <param name="depth">深度</param>
    public void Apply3DEffect(float depth = 10)
    {
        try
        {
            _shapeRange.ThreeD.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            _shapeRange.ThreeD.Depth = depth;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply 3D effect to shape range.", ex);
        }
    }

    /// <summary>
    /// 应用动画效果
    /// </summary>
    /// <param name="effectType">效果类型</param>
    /// <param name="triggerType">触发类型</param>
    public void ApplyAnimation(int effectType = 1, int triggerType = 1)
    {
        try
        {
            if (_shapeRange.AnimationSettings != null)
            {
                _shapeRange.AnimationSettings.EntryEffect = (MsPowerPoint.PpEntryEffect)effectType;
                _shapeRange.AnimationSettings.AdvanceMode = (MsPowerPoint.PpAdvanceMode)triggerType;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply animation to shape range.", ex);
        }
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _textFrame?.Dispose();
            _fill?.Dispose();
            _line?.Dispose();
            _shadow?.Dispose();
            _threeD?.Dispose();
            _tags?.Dispose();
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}

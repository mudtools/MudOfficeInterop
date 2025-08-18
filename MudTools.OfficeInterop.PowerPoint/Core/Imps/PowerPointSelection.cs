//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint Selection 对象的二次封装实现类
/// 实现 IPowerPointSelection 接口
/// </summary>
internal class PowerPointSelection : IPowerPointSelection
{
    private MsPowerPoint.Selection _selection;
    private bool _disposedValue = false;

    /// <summary>
    /// 初始化 PowerPointSelection 实例
    /// </summary>
    /// <param name="selection">要封装的 Microsoft.Office.Interop.PowerPoint.Selection 对象</param>
    internal PowerPointSelection(MsPowerPoint.Selection selection)
    {
        _selection = selection ?? throw new ArgumentNullException(nameof(selection));
    }

    #region 基础属性
    public object Parent => _selection.Parent;

    public IPowerPointApplication Application => _selection.Application != null ? new PowerPointApplication(_selection.Application) : null;

    public PpSelectionType Type => (PpSelectionType)_selection.Type;

    public int Count
    {
        get
        {
            try
            {
                switch (_selection.Type)
                {
                    case MsPowerPoint.PpSelectionType.ppSelectionNone:
                        return 0;
                    case MsPowerPoint.PpSelectionType.ppSelectionSlides:
                        return _selection.SlideRange?.Count ?? 0;
                    case MsPowerPoint.PpSelectionType.ppSelectionShapes:
                    case MsPowerPoint.PpSelectionType.ppSelectionText:
                        return _selection.ShapeRange?.Count ?? 0;
                }
            }
            catch { }
            return 0;
        }
    }
    #endregion

    #region 状态属性
    public bool IsEmpty => _selection.Type == MsPowerPoint.PpSelectionType.ppSelectionNone;

    #endregion

    #region 核心对象
    public IPowerPointShapeRange ShapeRange => _selection.ShapeRange != null ? new PowerPointShapeRange(_selection.ShapeRange) : null;

    public IPowerPointTextRange TextRange => _selection.TextRange != null ? new PowerPointTextRange(_selection.TextRange) : null;

    public IPowerPointSlideRange SlideRange => _selection.SlideRange != null ? new PowerPointSlideRange(_selection.SlideRange) : null;
    #endregion

    #region 操作方法
    public void Unselect()
    {
        _selection.Unselect();
    }

    public void SelectAll(bool replace = true)
    {
        try
        {
            if (Application?.ActiveWindow?.View is IPowerPointView slideView)
            {
                slideView.Slide.Shapes.Range(0).Select();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error selecting all: {ex.Message}");
        }
    }

    public void Copy()
    {
        _selection.Copy();
    }

    public void Cut()
    {
        _selection.Cut();
    }

    public void Delete()
    {
        _selection.Delete();
    }
    #endregion

    #region 文本操作
    public IPowerPointTextRange FindText(string findWhat, bool matchCase = false, bool matchWholeWord = false)
    {
        try
        {
            if (_selection.Type == MsPowerPoint.PpSelectionType.ppSelectionText && _selection.TextRange != null)
            {
                var foundRange = _selection.TextRange.Find(findWhat, 0,
                    matchCase ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                    matchWholeWord ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse);
                return foundRange != null ? new PowerPointTextRange(foundRange) : null;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error finding text: {ex.Message}");
        }
        return null;
    }

    public int ReplaceText(string findWhat, string replaceWhat, bool matchCase = false, bool matchWholeWord = false)
    {
        int replacements = 0;
        try
        {
            var foundRange = FindText(findWhat, matchCase, matchWholeWord);
            while (foundRange != null)
            {
                foundRange.Text = replaceWhat;
                replacements++;

                var foundRangeCom = _selection.TextRange.Find(findWhat,
                      0,
                      matchCase ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                      matchWholeWord ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse);

                foundRangeCom = null;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error replacing text: {ex.Message}");
        }
        return replacements;
    }

    public void SetTextFont(string fontName = "", float fontSize = 0, bool bold = false, bool italic = false)
    {
        try
        {
            if ((_selection.Type == MsPowerPoint.PpSelectionType.ppSelectionText || _selection.Type == MsPowerPoint.PpSelectionType.ppSelectionShapes) && _selection.TextRange != null)
            {
                var font = _selection.TextRange.Font;
                if (!string.IsNullOrEmpty(fontName))
                {
                    font.Name = fontName;
                }
                if (fontSize > 0)
                {
                    font.Size = fontSize;
                }
                font.Bold = bold ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
                font.Italic = italic ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error setting text font: {ex.Message}");
        }
    }

    public void SetTextColor(int color)
    {
        try
        {
            if ((_selection.Type == MsPowerPoint.PpSelectionType.ppSelectionText || _selection.Type == MsPowerPoint.PpSelectionType.ppSelectionShapes) && _selection.TextRange != null)
            {
                _selection.TextRange.Font.Color.RGB = color;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error setting text color: {ex.Message}");
        }
    }

    public void SetTextAlignment(PpParagraphAlignment alignment)
    {
        try
        {
            if ((_selection.Type == MsPowerPoint.PpSelectionType.ppSelectionText || _selection.Type == MsPowerPoint.PpSelectionType.ppSelectionShapes) && _selection.TextRange != null)
            {
                var parentShape = _selection.TextRange.Parent;
                _selection.TextRange.ParagraphFormat.Alignment = (MsPowerPoint.PpParagraphAlignment)alignment;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error setting text alignment: {ex.Message}");
        }
    }
    #endregion

    #region 形状操作
    public void AlignShapes(MsoAlignCmd alignment, int relativeTo = 0)
    {
        try
        {
            if ((_selection.Type == MsPowerPoint.PpSelectionType.ppSelectionShapes ||
                _selection.Type == MsPowerPoint.PpSelectionType.ppSelectionText) &&
                _selection.ShapeRange != null)
            {
                _selection.ShapeRange.Align((MsCore.MsoAlignCmd)alignment, (MsCore.MsoTriState)relativeTo);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error aligning shapes: {ex.Message}");
        }
    }

    public void DistributeShapes(MsoDistributeCmd distribution)
    {
        try
        {
            if ((_selection.Type == MsPowerPoint.PpSelectionType.ppSelectionShapes || _selection.Type == MsPowerPoint.PpSelectionType.ppSelectionText) && _selection.ShapeRange != null)
            {
                _selection.ShapeRange.Distribute((MsCore.MsoDistributeCmd)distribution, MsCore.MsoTriState.msoFalse);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error distributing shapes: {ex.Message}");
        }
    }

    public IPowerPointShape GroupShapes()
    {
        try
        {
            if ((_selection.Type == MsPowerPoint.PpSelectionType.ppSelectionShapes || _selection.Type == MsPowerPoint.PpSelectionType.ppSelectionText) && _selection.ShapeRange != null)
            {
                var groupedShape = _selection.ShapeRange.Group();
                return groupedShape != null ? new PowerPointShape(groupedShape) : null;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error grouping shapes: {ex.Message}");
        }
        return null;
    }

    public void UngroupShapes()
    {
        try
        {
            if ((_selection.Type == MsPowerPoint.PpSelectionType.ppSelectionShapes || _selection.Type == MsPowerPoint.PpSelectionType.ppSelectionText) && _selection.ShapeRange != null)
            {
                _selection.ShapeRange.Ungroup();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error ungrouping shapes: {ex.Message}");
        }
    }

    public void SetShapeFill(int color)
    {
        try
        {
            if ((_selection.Type == MsPowerPoint.PpSelectionType.ppSelectionShapes ||
                _selection.Type == MsPowerPoint.PpSelectionType.ppSelectionText) &&
                _selection.ShapeRange != null)
            {
                var shapes = _selection.ShapeRange;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    shapes[i].Fill.ForeColor.RGB = color;
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error setting shape fill: {ex.Message}");
        }
    }

    public void SetShapeBorder(int color, float weight = 1)
    {
        try
        {
            if ((_selection.Type == MsPowerPoint.PpSelectionType.ppSelectionShapes || _selection.Type == MsPowerPoint.PpSelectionType.ppSelectionText) && _selection.ShapeRange != null)
            {
                var shapes = _selection.ShapeRange;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    shapes[i].Line.ForeColor.RGB = color;
                    shapes[i].Line.Weight = weight;
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error setting shape border: {ex.Message}");
        }
    }
    #endregion

    #region 幻灯片操作
    public void CopySlides()
    {
        try
        {
            if (_selection.Type == MsPowerPoint.PpSelectionType.ppSelectionSlides && _selection.SlideRange != null)
            {
                _selection.SlideRange.Copy();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error copying slides: {ex.Message}");
        }
    }

    public void CutSlides()
    {
        try
        {
            if (_selection.Type == MsPowerPoint.PpSelectionType.ppSelectionSlides && _selection.SlideRange != null)
            {
                _selection.SlideRange.Cut();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error cutting slides: {ex.Message}");
        }
    }

    public void DeleteSlides()
    {
        try
        {
            if (_selection.Type == MsPowerPoint.PpSelectionType.ppSelectionSlides && _selection.SlideRange != null)
            {
                _selection.SlideRange.Delete();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error deleting slides: {ex.Message}");
        }
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            _disposedValue = true;
        }
    }

    ~PowerPointSelection()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}


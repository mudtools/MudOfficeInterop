//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint SlideRange 对象的二次封装实现类
/// 实现 IPowerPointSlideRange 接口
/// </summary>
internal class PowerPointSlideRange : IPowerPointSlideRange
{
    private MsPowerPoint.SlideRange _slideRange;
    private bool _disposedValue = false;

    /// <summary>
    /// 初始化 PowerPointSlideRange 实例
    /// </summary>
    /// <param name="slideRange">要封装的 Microsoft.Office.Interop.PowerPoint.SlideRange 对象</param>
    internal PowerPointSlideRange(MsPowerPoint.SlideRange slideRange)
    {
        _slideRange = slideRange ?? throw new ArgumentNullException(nameof(slideRange));
    }

    #region 基础属性
    public int Count => _slideRange.Count;

    public IPowerPointSlide this[int index] => new PowerPointSlide(_slideRange[index]);

    public object Parent => _slideRange.Parent;

    public IPowerPointApplication Application => _slideRange.Application != null ? new PowerPointApplication(_slideRange.Application) : null;

    public string Name
    {
        get => _slideRange.Name;
        set => _slideRange.Name = value;
    }
    #endregion


    #region 操作方法
    public void Select(bool replace = true)
    {
        _slideRange.Select();
    }

    public void Copy()
    {
        _slideRange.Copy();
    }

    public void Cut()
    {
        _slideRange.Cut();
    }

    public void Delete()
    {
        _slideRange.Delete();
    }

    public void Delete(int index)
    {
        try
        {
            var slideToDelete = _slideRange[index];
            slideToDelete?.Delete();
        }
        catch
        {
            // Handle error if index is invalid or deletion fails
        }
    }

    public void Delete(IPowerPointSlide slide)
    {
        if (slide is PowerPointSlide pptSlideWrapper)
        {
            try
            {
                pptSlideWrapper._slide?.Delete();
            }
            catch { /* Handle error */ }
        }
    }

    public void DeleteRange(int[] indices)
    {
        var sortedIndices = new List<int>(indices);
        sortedIndices.Sort((a, b) => b.CompareTo(a));
        foreach (int index in sortedIndices)
        {
            Delete(index);
        }
    }

    public void MoveTo(int toPos)
    {
        _slideRange.MoveTo(toPos);
    }

    public IPowerPointSlideRange Duplicate()
    {
        var duplicatedRange = _slideRange.Duplicate();
        return duplicatedRange != null ? new PowerPointSlideRange(duplicatedRange) : null;
    }
    #endregion

    #region 内容操作

    public IPowerPointSlide InsertNewSlide(int insertIndex, PpSlideLayout layout = PpSlideLayout.ppLayoutBlank)
    {
        try
        {
            var parentSlides = this.Parent as MsPowerPoint.Slides;
            if (parentSlides != null)
            {
                var newSlide = parentSlides.Add(insertIndex, (MsPowerPoint.PpSlideLayout)layout);
                return new PowerPointSlide(newSlide);
            }
            return null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error inserting new slide: {ex.Message}");
            return null;
        }
    }
    #endregion

    #region 导出和导入
    public bool ExportToPDF(string filename, bool overwrite = true)
    {
        try
        {
            if (System.IO.File.Exists(filename) && !overwrite)
            {
                return false;
            }
            _slideRange.Export(filename, "PDF");
            return true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error exporting SlideRange to PDF '{filename}': {ex.Message}");
            return false;
        }
    }

    public int ExportToImages(string folderPath, string format = "png", string prefix = "slide_")
    {
        if (!System.IO.Directory.Exists(folderPath)) return 0;

        int count = 0;
        try
        {
            for (int i = 1; i <= _slideRange.Count; i++)
            {
                string fileName = System.IO.Path.Combine(folderPath, $"{prefix}{i}.{format}");
                _slideRange.Export(fileName, format.ToUpper());
                if (System.IO.File.Exists(fileName))
                {
                    count++;
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error exporting SlideRange to images: {ex.Message}");
        }
        return count;
    }
    #endregion

    #region IEnumerable<IPowerPointSlide> Support
    public IEnumerator<IPowerPointSlide> GetEnumerator()
    {
        for (int i = 1; i <= _slideRange.Count; i++)
        {
            yield return new PowerPointSlide(_slideRange[i]);
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (_slideRange != null)
            {
                try
                {
                    Marshal.FinalReleaseComObject(_slideRange);
                }
                catch
                {
                    // 忽略释放过程中可能发生的异常
                }
                _slideRange = null;
            }
            _slideRange = null;

            _disposedValue = true;
        }
    }

    ~PowerPointSlideRange()
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

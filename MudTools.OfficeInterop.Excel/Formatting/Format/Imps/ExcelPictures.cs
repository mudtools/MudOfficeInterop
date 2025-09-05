//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Pictures 集合对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Pictures 对象的安全访问和资源管理
/// </summary>
internal class ExcelPictures : IExcelPictures
{
    /// <summary>
    /// 底层的 COM Pictures 集合对象
    /// </summary>
    private MsExcel.Pictures _pictures;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelPictures 实例
    /// </summary>
    /// <param name="pictures">底层的 COM Pictures 集合对象</param>
    internal ExcelPictures(MsExcel.Pictures pictures)
    {
        _pictures = pictures;
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放所有子图片对象
                for (int i = 1; i <= Count; i++)
                {
                    var picture = this[i] as ExcelPicture;
                    picture?.Dispose();
                }

                // 释放底层COM对象
                if (_pictures != null)
                    Marshal.ReleaseComObject(_pictures);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _pictures = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    /// <summary>
    /// 获取图片集合中的图片数量
    /// </summary>
    public int Count => _pictures?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的图片对象
    /// </summary>
    /// <param name="index">图片索引（从1开始）</param>
    /// <returns>图片对象</returns>
    public IExcelPicture? this[int index]
    {
        get
        {
            if (_pictures == null || index < 1 || index > Count)
                return null;

            try
            {
                MsExcel.Picture? picture = _pictures.Item(index) as MsExcel.Picture;
                return picture != null ? new ExcelPicture(picture) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取指定名称的图片对象
    /// </summary>
    /// <param name="name">图片名称</param>
    /// <returns>图片对象</returns>
    public IExcelPicture? this[string name]
    {
        get
        {
            if (_pictures == null || string.IsNullOrEmpty(name))
                return null;

            try
            {
                var picture = _pictures.Item(name) as MsExcel.Picture;
                return picture != null ? new ExcelPicture(picture) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取图片集合所在的父对象
    /// </summary>
    public object Parent => _pictures?.Parent;

    #endregion

    #region 创建和添加

    /// <summary>
    /// 向工作表添加图片
    /// </summary>
    /// <param name="filename">图片文件路径</param>
    /// <param name="Converter"></param>
    /// <returns>新创建的图片对象</returns>
    public IExcelPicture? Insert(string filename, object Converter)
    {
        if (_pictures == null || string.IsNullOrEmpty(filename))
            return null;

        try
        {
            // 验证文件是否存在
            if (!File.Exists(filename))
                return null;

            var picture = _pictures.Insert(filename, Converter);

            return picture != null ? new ExcelPicture(picture) : null;
        }
        catch
        {
            return null;
        }
    }


    /// <summary>
    /// 从字节数组添加图片
    /// </summary>
    /// <param name="imageBytes">图片字节数组</param>
    /// <param name="imageFormat">图片格式</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的图片对象</returns>
    public IExcelPicture AddFromBytes(byte[] imageBytes, string imageFormat = "png",
                                    double left = 0, double top = 0, double width = -1, double height = -1)
    {
        if (_pictures == null || imageBytes == null || imageBytes.Length == 0)
            return null;

        try
        {
            // 创建临时文件
            string tempPath = Path.GetTempFileName();
            string tempImagePath = Path.ChangeExtension(tempPath, imageFormat);

            // 重命名临时文件
            File.Move(tempPath, tempImagePath);

            // 写入图片数据
            File.WriteAllBytes(tempImagePath, imageBytes);

            // 添加图片
            var picture = Insert(tempImagePath, null);

            // 删除临时文件
            if (File.Exists(tempImagePath))
                File.Delete(tempImagePath);

            return picture;
        }
        catch
        {
            return null;
        }
    }

    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据名称查找图片
    /// </summary>
    /// <param name="name">图片名称</param>
    /// <returns>匹配的图片数组</returns>
    public IExcelPicture[] FindByName(string name)
    {
        if (_pictures == null || string.IsNullOrEmpty(name) || Count == 0)
            return [];

        var result = new List<IExcelPicture>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var picture = this[i];
                if (picture != null && picture.Name?.Contains(name) == true)
                {
                    result.Add(picture);
                }
            }
            catch
            {
                // 忽略单个图片访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据位置查找图片
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的图片数组</returns>
    public IExcelPicture[] FindByPosition(double left, double top, double tolerance = 10)
    {
        if (_pictures == null || Count == 0)
            return [];

        var result = new List<IExcelPicture>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var picture = this[i];
                if (picture != null)
                {
                    double picLeft = picture.Left;
                    double picTop = picture.Top;

                    if (Math.Abs(picLeft - left) <= tolerance && Math.Abs(picTop - top) <= tolerance)
                    {
                        result.Add(picture);
                    }
                }
            }
            catch
            {
                // 忽略单个图片访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据大小查找图片
    /// </summary>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的图片数组</returns>
    public IExcelPicture[] FindBySize(double width, double height, double tolerance = 10)
    {
        if (_pictures == null || Count == 0)
            return [];

        var result = new List<IExcelPicture>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var picture = this[i];
                if (picture != null)
                {
                    double picWidth = picture.Width;
                    double picHeight = picture.Height;

                    if (Math.Abs(picWidth - width) <= tolerance && Math.Abs(picHeight - height) <= tolerance)
                    {
                        result.Add(picture);
                    }
                }
            }
            catch
            {
                // 忽略单个图片访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取指定区域内的所有图片
    /// </summary>
    /// <param name="range">目标区域</param>
    /// <returns>区域内的图片数组</returns>
    public IExcelPicture[] GetPicturesInRange(IExcelRange range)
    {
        if (_pictures == null || range == null || Count == 0)
            return [];

        var result = new List<IExcelPicture>();
        // 注意：Excel Pictures集合不直接支持区域筛选
        // 这里返回所有图片作为示例
        for (int i = 1; i <= Count; i++)
        {
            var picture = this[i];
            if (picture != null)
                result.Add(picture);
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取可见的图片
    /// </summary>
    /// <returns>可见图片数组</returns>
    public IExcelPicture[] GetVisiblePictures()
    {
        if (_pictures == null || Count == 0)
            return [];

        var result = new List<IExcelPicture>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var picture = this[i];
                if (picture != null && picture.Visible)
                {
                    result.Add(picture);
                }
            }
            catch
            {
                // 忽略单个图片访问异常
            }
        }
        return result.ToArray();
    }

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除所有图片
    /// </summary>
    public void Clear()
    {
        if (_pictures == null) return;

        try
        {
            // 从后往前删除，避免索引变化问题
            for (int i = Count; i >= 1; i--)
            {
                try
                {
                    ((MsExcel.Picture)_pictures.Item(i)).Delete();
                }
                catch
                {
                    // 忽略删除过程中的异常
                }
            }
        }
        catch
        {
            // 忽略清空过程中的异常
        }
    }

    /// <summary>
    /// 删除指定索引的图片
    /// </summary>
    /// <param name="index">要删除的图片索引</param>
    public void Delete(int index)
    {
        if (_pictures == null || index < 1 || index > Count)
            return;

        try
        {
            ((MsExcel.Picture)_pictures.Item(index)).Delete();
        }
        catch
        {
            // 忽略删除过程中的异常
        }
    }

    /// <summary>
    /// 删除指定的图片
    /// </summary>
    /// <param name="picture">要删除的图片对象</param>
    public void Delete(IExcelPicture picture)
    {
        if (_pictures == null || picture == null)
            return;

        try
        {
            picture.Delete();
        }
        catch
        {
            // 忽略删除过程中的异常
        }
    }

    /// <summary>
    /// 批量删除图片
    /// </summary>
    /// <param name="indices">要删除的图片索引数组</param>
    public void DeleteRange(int[] indices)
    {
        if (_pictures == null || indices == null || indices.Length == 0)
            return;

        // 按降序排列索引，避免删除时索引变化
        Array.Sort(indices, (a, b) => b.CompareTo(a));

        foreach (int index in indices)
        {
            Delete(index);
        }
    }


    #endregion


    public IEnumerator<IExcelPicture> GetEnumerator()
    {
        for (int i = 0; i < Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
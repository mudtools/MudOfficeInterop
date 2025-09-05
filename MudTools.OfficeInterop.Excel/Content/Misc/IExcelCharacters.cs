//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel中的字符对象接口，提供对单元格中文本的字符级操作功能
/// </summary>
public interface IExcelCharacters : IDisposable
{
    /// <summary>
    /// 获取条件值对象所在的Application对象
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取字符对象的父级对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取字符数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取或设置文本内容
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取字符的字体属性
    /// </summary>
    IExcelFont Font { get; }

    /// <summary>
    /// 删除字符
    /// </summary>
    void Delete();


    IExcelCharacters this[int? start, int? length] { get; }

    /// <summary>
    /// 插入文本到指定位置
    /// </summary>
    /// <param name="text">要插入的文本</param>
    /// <returns>插入后的字符对象</returns>
    IExcelCharacters Insert(string text);

    /// <summary>
    /// 查找文本在字符中的位置
    /// </summary>
    /// <param name="what">要查找的文本</param>
    /// <param name="after">起始查找位置</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="matchWholeWord">是否匹配整个单词</param>
    /// <returns>找到的位置，未找到返回0</returns>
    int Find(string what, int after = 1, bool matchCase = false, bool matchWholeWord = false);

    /// <summary>
    /// 替换文本
    /// </summary>
    /// <param name="what">要替换的文本</param>
    /// <param name="replacement">替换文本</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="matchWholeWord">是否匹配整个单词</param>
    /// <returns>替换的次数</returns>
    int Replace(string what, string replacement, bool matchCase = false, bool matchWholeWord = false);

}

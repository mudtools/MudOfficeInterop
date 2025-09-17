//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel工作表中的绘图对象集合，提供对图表、形状等绘图对象的管理功能。
/// 该接口继承自IDisposable和IEnumerable&lt;IExcelDrawing&gt;，支持资源释放和遍历操作。
/// </summary>
public interface IExcelDrawingObjects : IExcelComGraphObjects, IEnumerable<IExcelDrawing>, IDisposable
{
    /// <summary>
    /// 获取绘图对象的堆叠次序
    /// </summary>
    int ZOrder { get; }

    /// <summary>
    /// 获取或设置滚动条或数值调节按钮的最小值
    /// </summary>
    int Min { get; set; }

    /// <summary>
    /// 获取或设置滚动条或数值调节按钮的最大值
    /// </summary>
    int Max { get; set; }

    /// <summary>
    /// 获取或设置单击滚动条上的滚动框与箭头之间的区域时，滚动条滚动的幅度
    /// </summary>
    int SmallChange { get; set; }

    /// <summary>
    /// 获取或设置下拉列表框中显示的列表项数目
    /// </summary>
    int DropDownLines { get; set; }

    /// <summary>
    /// 获取或设置控件是否具有默认的命令按钮功能
    /// </summary>
    bool DefaultButton { get; set; }

    /// <summary>
    /// 获取或设置当用户单击控件时，是否关闭对话框
    /// </summary>
    bool DismissButton { get; set; }

    /// <summary>
    /// 获取或设置是否显示垂直滚动条
    /// </summary>
    bool DisplayVerticalScrollBar { get; set; }

    /// <summary>
    /// 获取或设置控件是否使用三维阴影效果
    /// </summary>
    bool Display3DShading { get; set; }

    /// <summary>
    /// 获取或设置文本是否自动缩进
    /// </summary>
    bool AddIndent { get; set; }

    /// <summary>
    /// 获取或设置控件中文本的阅读次序
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取或设置控件中显示的文本
    /// </summary>
    string Text { get; set; }


    /// <summary>
    /// 获取或设置列表框中当前选中项的索引
    /// </summary>
    int ListIndex { get; set; }

    /// <summary>
    /// 获取或设置输入框的类型
    /// </summary>
    int InputType { get; set; }

    /// <summary>
    /// 获取或设置与控件关联的工作表区域
    /// </summary>
    string LinkedCell { get; set; }

    /// <summary>
    /// 获取或设置单击滚动条箭头时滚动条滚动的幅度
    /// </summary>
    int LargeChange { get; set; }

    /// <summary>
    /// 获取或设置列表框填充区域的地址
    /// </summary>
    string ListFillRange { get; set; }

    /// <summary>
    /// 获取或设置控件的当前值
    /// </summary>
    int Value { get; set; }


    /// <summary>
    /// 获取或设置控件的标题
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// 获取或设置控件中的文本是否被锁定
    /// </summary>
    bool LockedText { get; set; }

    /// <summary>
    /// 获取或设置控件是否支持多行文本
    /// </summary>
    bool MultiLine { get; set; }

    /// <summary>
    /// 获取或设置列表框是否支持多项选择
    /// </summary>
    bool MultiSelect { get; set; }

    /// <summary>
    /// 获取或设置控件是否具有圆角
    /// </summary>
    bool RoundedCorners { get; set; }

    /// <summary>
    /// 获取或设置控件是否可用
    /// </summary>
    bool Enabled { get; set; }



    /// <summary>
    /// 获取控件的字符属性
    /// </summary>
    IExcelCharacters? Characters { get; }

    /// <summary>
    /// 获取控件的字体属性
    /// </summary>
    IExcelFont? Font { get; }


    /// <summary>
    /// 根据索引获取绘图对象（索引从1开始）
    /// </summary>
    /// <param name="index">绘图对象索引</param>
    /// <returns>绘图对象</returns>
    IExcelDrawing? this[int index] { get; }

    /// <summary>
    /// 根据名称获取绘图对象
    /// </summary>
    /// <param name="name">绘图对象名称</param>
    /// <returns>绘图对象</returns>
    IExcelDrawing? this[string name] { get; }

    /// <summary>
    /// 根据索引或名称获取绘图对象
    /// </summary>
    /// <param name="index">绘图对象索引或名称</param>
    /// <returns>绘图对象</returns>
    IExcelDrawing GetItem(object index);

    /// <summary>
    /// 根据名称查找绘图对象
    /// </summary>
    /// <param name="name">对象名称</param>
    /// <returns>绘图对象</returns>
    IExcelDrawing FindByName(string name);

    /// <summary>
    /// 删除指定名称的绘图对象
    /// </summary>
    /// <param name="name">对象名称</param>
    void Remove(string name);

    /// <summary>
    /// 清除所有绘图对象
    /// </summary>
    void Clear();

    /// <summary>
    /// 获取可见的绘图对象
    /// </summary>
    IEnumerable<IExcelDrawing> VisibleItems { get; }

    /// <summary>
    /// 获取锁定的绘图对象
    /// </summary>
    IEnumerable<IExcelDrawing> LockedItems { get; }
}
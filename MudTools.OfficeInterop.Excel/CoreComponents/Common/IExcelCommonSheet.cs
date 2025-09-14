//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Sheet公共接口，Excel的Sheet类型：WordSheet、Chart。
/// </summary>
public interface IExcelCommonSheet : IDisposable
{
    /// <summary>
    /// 获取图表所在的 Excel Application 对象
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置工作表的名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取工作表类型
    /// </summary>
    XlSheetType Type { get; }

    /// <summary>
    /// 获取图表对象的索引位置
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取工作表的超链接集合
    /// </summary>
    IExcelHyperlinks Hyperlinks { get; }

    /// <summary>
    /// 获取工作表所在的父对象（通常是工作簿）
    /// </summary>
    object? Parent { get; }

    /// <summary>
    ///  获取工作表所在的父对象名字。
    /// </summary>
    string? ParentName { get; }


    /// <summary>
    /// 获取工作表所在的父对象（通常是工作簿）
    /// 对应 Worksheet.Parent 属性
    /// </summary>
    IExcelWorkbook? ParentWorkbook { get; }

    /// <summary>
    /// 获取图表是否被保护
    /// </summary>
    bool IsProtected { get; }


    /// <summary>
    /// 获取一个值，该值指示工作表当前是否处于保护模式
    /// </summary>
    bool ProtectionMode { get; }

    /// <summary>
    /// 获取或设置工作表是否可见
    /// </summary>
    bool IsVisible { get; set; }

    /// <summary>
    /// 获取或设置图表是否可见
    /// </summary>
    XlSheetVisibility Visible { get; set; }

    /// <summary>
    /// 获取工作表的页面设置对象
    /// </summary>
    IExcelPageSetup PageSetup { get; }

    /// <summary>
    /// 获取工作表的形状集合
    /// </summary>
    IExcelShapes? Shapes { get; }

    /// <summary>
    /// 获取工作表内容是否受保护
    /// </summary>
    bool ProtectContents { get; }

    /// <summary>
    /// 将工作表另存为xlsx文件。
    /// </summary>
    /// <param name="filePath"></param>
    void SaveAs(string filePath);

    /// <summary>
    /// 打印预览
    /// </summary>
    void PrintPreview();

    /// <summary>
    /// 打印工作表
    /// </summary>
    /// <param name="preview">是否打印预览</param>
    void PrintOut(bool preview = false);

    /// <summary>
    /// 删除工作表
    /// </summary>
    void Delete();

    /// <summary>
    /// 激活对象
    /// </summary>
    void Activate();

    /// <summary>
    /// 选择对象
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 复制对象
    /// </summary>
    void Copy();

    /// <summary>
    /// 复制工作表
    /// 对应 Worksheet.Copy 方法
    /// </summary>
    /// <param name="before">复制到指定工作表之前</param>
    /// <param name="after">复制到指定工作表之后</param>
    void Copy(IExcelCommonSheet? before = null, IExcelCommonSheet? after = null);

    /// <summary>
    /// 移动工作表
    /// 对应 Worksheet.Move 方法
    /// </summary>
    /// <param name="before">移动到指定工作表之前</param>
    /// <param name="after">移动到指定工作表之后</param>
    void Move(IExcelCommonSheet? before = null, IExcelCommonSheet? after = null);

    /// <summary>
    /// 获取工作表中的OLE对象集合或指定索引的OLE对象
    /// </summary>
    /// <param name="index">OLE对象的索引，如果为null则返回所有OLE对象集合</param>
    /// <returns>OLE对象或OLE对象集合</returns>
    object? OLEObjects(int? index = null);

    /// <summary>
    /// 清除图表内容
    /// </summary>
    void Clear();


    /// <summary>
    /// 清除工作表内容和格式
    /// </summary>
    void ClearAll();


    /// <summary>
    /// 清除工作表内容
    /// </summary>
    void ClearContents();

    /// <summary>
    /// 取消保护工作表
    /// 对应 Worksheet.Unprotect 方法
    /// </summary>
    /// <param name="password">保护密码</param>
    void Unprotect(string password = "");

    /// <summary>
    /// 保护工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    /// <param name="drawingObjects">是否保护图形对象</param>
    /// <param name="contents">是否保护内容</param>
    /// <param name="scenarios">是否保护方案</param>
    /// <param name="userInterfaceOnly">是否仅保护用户界面</param>
    public void Protect(string? password = null, bool? drawingObjects = null, bool? contents = null, bool? scenarios = null, bool? userInterfaceOnly = null);

}
//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;
/// <summary>
/// VBE VBComponent 对象的二次封装接口
/// 提供对 Microsoft.Vbe.Interop.VBComponent 的安全访问和操作
/// </summary>
public interface IVbeVBComponent : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取 VB 组件的名称
    /// 对应 VBComponent.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取 VB 组件的类型
    /// 对应 VBComponent.Type 属性
    /// </summary>
    vbext_ComponentType Type { get; }

    /// <summary>
    /// 获取 VB 组件所在的Application对象（VBE 对象）
    /// 对应 VBComponent.Application 属性
    /// </summary>
    IVbeApplication Application { get; }

    /// <summary>
    /// 获取 VB 组件的代码模块对象
    /// 对应 VBComponent.CodeModule 属性
    /// </summary>
    IVbeCodeModule CodeModule { get; }

    /// <summary>
    /// 获取 VB 组件的设计时对象（例如，UserForm 的设计器）
    /// 对应 VBComponent.Designer 属性
    /// </summary>
    object Designer { get; } // 使用 object 作为通用占位符

    /// <summary>
    /// 获取 VB 组件的设计器数据（特定于设计器的二进制数据）
    /// 对应 VBComponent.DesignerID 属性 (概念上接近)
    /// </summary>
    string DesignerID { get; }

    /// <summary>
    /// 获取 VB 组件是否已保存
    /// </summary>
    bool IsSaved { get; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 选择 VB 组件
    /// 对应 VBComponent.Activate 方法
    /// </summary>
    void Activate();

    /// <summary>
    /// 导出 VB 组件到文件
    /// 对应 VBComponent.Export 方法
    /// </summary>
    /// <param name="fileName">导出文件路径</param>
    void Export(string fileName);

    #endregion

    #region 代码操作   

    /// <summary>
    /// 清除 VB 组件中的所有代码
    /// </summary>
    void ClearCode();
    #endregion

    #region 格式设置
    /// <summary>
    /// 设置 VB 组件属性
    /// </summary>
    /// <param name="propertyName">属性名称</param>
    /// <param name="value">属性值</param>
    void SetProperty(string propertyName, object value);

    /// <summary>
    /// 获取 VB 组件属性
    /// </summary>
    /// <param name="propertyName">属性名称</param>
    /// <returns>属性值</returns>
    object GetProperty(string propertyName);
    #endregion

    #region 导出和转换
    /// <summary>
    /// 获取 VB 组件的代码文本
    /// </summary>
    /// <returns>代码文本</returns>
    string GetCodeText();

    /// <summary>
    /// 设置 VB 组件的代码文本
    /// </summary>
    /// <param name="codeText">新的代码文本</param>
    void SetCodeText(string codeText);

    /// <summary>
    /// 将 VB 组件转换为字符串表示（例如，包含名称和代码）
    /// </summary>
    /// <returns>字符串表示</returns>
    string ToString();

    /// <summary>
    /// 获取 VB 组件的字节数据（如果适用，例如二进制形式）
    /// </summary>
    /// <returns>字节数组</returns>
    byte[] GetBytes();
    #endregion

}

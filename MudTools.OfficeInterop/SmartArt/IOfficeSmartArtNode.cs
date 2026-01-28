//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// SmartArtNode 封装接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeSmartArtNode : IOfficeObject<IOfficeSmartArtNode, MsCore.SmartArtNode>, IDisposable
{
    /// <summary>
    /// 获取节点的父对象
    /// </summary>
    /// <value>
    /// 返回包含此节点的父对象，如果此节点是根节点则可能为null
    /// </value>
    object? Parent { get; }

    /// <summary>
    /// 获取关联的 Shape 对象（如果存在）
    /// </summary>
    IOfficeShapeRange? Shapes { get; }

    /// <summary>
    /// 获取与该 SmartArt 节点关联的文本框对象
    /// </summary>
    /// <value>
    /// 返回一个 IOfficeTextFrame2 对象，用于访问和操作与此 SmartArt 节点关联的文本内容和格式设置
    /// </value>
    IOfficeTextFrame2? TextFrame2 { get; }

    /// <summary>
    /// 获取子节点集合
    /// </summary>
    IOfficeSmartArtNodes? Nodes { get; }

    /// <summary>
    /// 获取父节点
    /// </summary>
    IOfficeSmartArtNode? ParentNode { get; }

    /// <summary>
    /// 获取节点的类型
    /// </summary>
    MsoSmartArtNodeType Type { get; }


    /// <summary>
    /// 获取节点所在的层级
    /// </summary>
    int Level { get; }

    /// <summary>
    /// 获取节点是否为隐藏状态
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Hidden { get; }

    /// <summary>
    /// 删除当前节点及其所有子节点
    /// </summary>
    void Delete();

    /// <summary>
    /// 在当前节点中添加一个新节点
    /// </summary>
    /// <param name="Position">指定新节点相对于当前节点的位置，默认为默认位置</param>
    /// <param name="Type">指定新节点的类型，默认为默认类型</param>
    /// <returns>返回新创建的节点对象，如果添加失败则返回null</returns>
    IOfficeSmartArtNode? AddNode(
          MsoSmartArtNodePosition Position = MsoSmartArtNodePosition.msoSmartArtNodeDefault,
          MsoSmartArtNodeType Type = MsoSmartArtNodeType.msoSmartArtNodeTypeDefault);

    /// <summary>
    /// 增大当前节点的尺寸
    /// </summary>
    void Larger();

    /// <summary>
    /// 减小当前节点的尺寸
    /// </summary>
    void Smaller();

    /// <summary>
    /// 将当前节点在节点列表中向上移动一位
    /// </summary>
    void ReorderUp();

    /// <summary>
    /// 将当前节点在节点列表中向下移动一位
    /// </summary>
    void ReorderDown();

    /// <summary>
    /// 提升节点层级（向根靠近）
    /// </summary>
    void Promote();

    /// <summary>
    /// 降低节点层级（向叶靠近）
    /// </summary>
    void Demote();
}
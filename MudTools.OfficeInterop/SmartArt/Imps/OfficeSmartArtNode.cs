//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Imps;


// SmartArtNode 实现类
internal class OfficeSmartArtNode : IOfficeSmartArtNode
{
    private static readonly ILog log = LogManager.GetLogger(typeof(OfficeSmartArtNode));

    internal MsCore.SmartArtNode? _node;
    private bool _disposedValue;

    internal OfficeSmartArtNode(MsCore.SmartArtNode node)
    {
        _node = node ?? throw new ArgumentNullException(nameof(node));
        _disposedValue = false;
    }

    #region 属性实现    

    /// <summary>
    /// 获取或设置节点文本内容
    /// </summary>
    public string Text
    {
        get
        {
            if (_node == null || _node.TextFrame2 == null || _node.TextFrame2.TextRange == null)
                return string.Empty;
            try
            {
                return _node.TextFrame2.TextRange.Text;
            }
            catch (Exception x)
            {
                log.Warn($"读取节点文本失败: {x.Message}");
                return string.Empty;
            }
        }
        set
        {
            if (_node == null || _node.TextFrame2 == null || _node.TextFrame2.TextRange == null) return;
            try
            {
                _node.TextFrame2.TextRange.Text = value;
            }
            catch (Exception x)
            {
                log.Error($"设置节点文本失败: {x.Message}");
            }
        }
    }

    /// <summary>
    /// 获取关联的 Shape 对象
    /// </summary>
    public IOfficeShapeRange? Shapes
    {
        get
        {
            if (_node == null || _node.Shapes == null) return null;
            try
            {
                return new OfficeShapeRange(_node.Shapes);
            }
            catch (Exception x)
            {
                log.Warn($"获取关联形状失败: {x.Message}");
                return null;
            }
        }
    }


    /// <summary>
    /// 获取父节点
    /// </summary>
    public object? Parent
    {
        get
        {
            if (_node == null || _node.Parent == null) return null;
            return _node.Parent;
        }
    }

    /// <summary>
    /// 获取子节点集合
    /// </summary>
    public IOfficeSmartArtNodes Nodes
    {
        get
        {
            if (_node == null || _node.Nodes == null) return null;
            try
            {
                return new OfficeSmartArtNodes(_node.Nodes);
            }
            catch (Exception x)
            {
                log.Warn($"获取子节点集合失败: {x.Message}");
                return null;
            }
        }
    }

    public IOfficeSmartArtNode ParentNode
    {
        get
        {
            if (_node == null || _node.ParentNode == null) return null;
            try
            {
                return new OfficeSmartArtNode(_node.ParentNode);
            }
            catch (Exception x)
            {
                log.Warn($"获取父节点失败: {x.Message}");
                return null;
            }
        }
    }

    public MsoSmartArtNodeType Type
    {
        get
        {
            if (_node == null) return MsoSmartArtNodeType.msoSmartArtNodeTypeDefault;
            return _node.Type.EnumConvert(MsoSmartArtNodeType.msoSmartArtNodeTypeDefault);
        }
    }

    public int Level
    {
        get
        {
            if (_node == null) return 0;
            return _node.Level;
        }
    }

    public bool Hidden
    {
        get
        {
            if (_node == null) return false;
            return _node.Hidden.ConvertToBool();
        }
    }

    /// <summary>
    /// 判断是否为根节点（无父节点）
    /// </summary>
    public bool IsRoot => Parent == null;

    #endregion

    #region 方法实现

    /// <summary>
    /// 删除当前节点及其所有子节点
    /// </summary>
    public void Delete()
    {
        if (_node == null) return;
        try
        {
            _node.Delete();
            _node = null; // 避免重复操作
        }
        catch (Exception x)
        {
            log.Error($"删除 SmartArt 节点失败: {x.Message}");
        }
    }


    public IOfficeSmartArtNode? AddNode(
        MsoSmartArtNodePosition Position = MsoSmartArtNodePosition.msoSmartArtNodeDefault,
        MsoSmartArtNodeType Type = MsoSmartArtNodeType.msoSmartArtNodeTypeDefault)
    {
        if (_node == null) return null;
        try
        {
            var newNode = _node.AddNode(Position.EnumConvert(MsCore.MsoSmartArtNodePosition.msoSmartArtNodeDefault),
                             Type.EnumConvert(MsCore.MsoSmartArtNodeType.msoSmartArtNodeTypeDefault));
            if (newNode != null)
            {
                return new OfficeSmartArtNode(newNode);
            }
            return null;
        }
        catch (Exception x)
        {
            log.Error($"添加子节点失败: {x.Message}");
            return null;
        }
    }

    /// <summary>
    /// 提升节点层级（向根靠近）
    /// </summary>
    public void Promote()
    {
        if (_node == null) return;
        try
        {
            _node.Promote();
        }
        catch (Exception x)
        {
            log.Error($"提升节点失败: {x.Message}");
        }
    }

    /// <summary>
    /// 降低节点层级（向叶靠近）
    /// </summary>
    public void Demote()
    {
        if (_node == null) return;
        try
        {
            _node.Demote();
        }
        catch (Exception x)
        {
            log.Error($"降低节点失败: {x.Message}");
        }
    }

    public void Larger()
    {
        if (_node == null) return;
        try
        {
            _node.Larger();
        }
        catch (Exception x)
        {
            log.Error($"节点变大失败: {x.Message}");
        }
    }

    public void Smaller()
    {
        if (_node == null) return;
        try
        {
            _node.Smaller();
        }
        catch (Exception x)
        {
            log.Error($"节点变小失败: {x.Message}");
        }
    }

    public void ReorderUp()
    {
        if (_node == null) return;
        try
        {
            _node.ReorderUp();
        }
        catch (Exception x)
        {
            log.Error($"节点上移失败: {x.Message}");
        }
    }

    public void ReorderDown()
    {
        if (_node == null) return;
        try
        {
            _node.ReorderDown();
        }
        catch (Exception x)
        {
            log.Error($"节点下移失败: {x.Message}");
        }
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _node != null)
        {
            Marshal.ReleaseComObject(_node);
            _node = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    ~OfficeSmartArtNode()
    {
        Dispose(false);
    }

    #endregion
}
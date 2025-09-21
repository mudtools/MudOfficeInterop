//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Imps;


// SmartArtNodes 集合实现类
internal class OfficeSmartArtNodes : IOfficeSmartArtNodes
{
    private static readonly ILog log = LogManager.GetLogger(typeof(OfficeSmartArtNodes));

    private readonly MsCore.SmartArtNodes? _nodes;
    private bool _disposedValue;

    internal OfficeSmartArtNodes(MsCore.SmartArtNodes nodes)
    {
        _nodes = nodes;
        _disposedValue = false;
    }
    #region 属性与索引器

    /// <summary>
    /// 获取节点总数
    /// </summary>
    public int Count => _nodes?.Count ?? 0;

    /// <summary>
    /// 根据索引获取节点（索引从1开始）
    /// </summary>
    public IOfficeSmartArtNode? this[int index]
    {
        get
        {
            if (_nodes == null || index < 1 || index > Count) return null;
            try
            {
                var comNode = _nodes[index];
                return new OfficeSmartArtNode(comNode);
            }
            catch (Exception x)
            {
                log.Warn($"获取第 {index} 个节点失败: {x.Message}");
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <summary>
    /// 在集合末尾添加一个新节点
    /// </summary>
    public IOfficeSmartArtNode? Add(string text)
    {
        if (_nodes == null) return null;
        try
        {
            var newNode = _nodes.Add();
            if (newNode.TextFrame2?.TextRange != null)
            {
                newNode.TextFrame2.TextRange.Text = text;
            }
            return new OfficeSmartArtNode(newNode);
        }
        catch (Exception x)
        {
            log.Error($"添加 SmartArt 节点失败: {x.Message}");
            return null;
        }
    }

    /// <summary>
    /// 清空所有节点（注意：某些 SmartArt 布局至少需保留一个节点）
    /// </summary>
    public void Clear()
    {
        if (_nodes == null) return;
        try
        {
            while (_nodes.Count > 1)
            {
                _nodes[1].Delete();
            }
            if (_nodes.Count == 1)
            {
                _nodes[1].TextFrame2?.TextRange?.Delete();
            }
        }
        catch (Exception x)
        {
            log.Error($"清空 SmartArt 节点失败: {x.Message}");
        }
    }

    #endregion

    #region IEnumerable 实现

    public IEnumerator<IOfficeSmartArtNode> GetEnumerator()
    {
        if (_nodes == null) yield break;

        for (int i = 1; i <= _nodes.Count; i++)
        {
            var comNode = _nodes[i];
            yield return new OfficeSmartArtNode(comNode);
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _nodes != null)
        {
            Marshal.ReleaseComObject(_nodes);
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    ~OfficeSmartArtNodes()
    {
        Dispose(false);
    }

    #endregion
}
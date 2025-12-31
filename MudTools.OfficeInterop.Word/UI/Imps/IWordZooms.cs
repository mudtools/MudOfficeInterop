
namespace MudTools.OfficeInterop.Word.Imps;

partial class WordZooms
{
    public int Count
    {
        get
        {
            return 5;
        }
    }

    public IWordZoom? this[int index]
    {
        get
        {
            var viewType = index.EnumConvert(WdViewType.wdMasterView);
            var zoom = this[viewType];
            return zoom;
        }
    }
}
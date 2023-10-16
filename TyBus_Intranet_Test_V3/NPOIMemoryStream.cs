using System.IO;

namespace TyBus_Intranet_Test_V3
{
    public class NPOIMemoryStream : MemoryStream
    {
        public NPOIMemoryStream()
        {
            AllowClose = true;
        }

        public bool AllowClose { get; set; }

        public override void Close()
        {
            if (AllowClose)
                base.Close();
        }
    }
}
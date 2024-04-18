using System.Drawing;
using System.Windows.Forms;
using AccessCodeLib.AccUnit.VbeAddIn.Resources;
using AccessCodeLib.Common.Tools.Logging;
using stdole;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    /// @see http://blogs.msdn.com/b/andreww/archive/2007/07/30/converting-between-ipicturedisp-and-system-drawing-image.aspx
    public class AxHostConverter : AxHost
    {
        private AxHostConverter() : base("") { }

        public static IPictureDisp ImageToPictureDisp(Image image)
        {
            using (new BlockLogger())
            {
                return (IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }

        public static Image PictureDispToImage(IPictureDisp pictureDisp)
        {
            return GetPictureFromIPicture(pictureDisp);
        }

        private static IPictureDisp _runTestPictureDisp;
        public static IPictureDisp RunTestPictureDisp
        {
            get { return _runTestPictureDisp ?? (_runTestPictureDisp = ImageToPictureDisp(Icons.runtest)); }
        }

    }
}

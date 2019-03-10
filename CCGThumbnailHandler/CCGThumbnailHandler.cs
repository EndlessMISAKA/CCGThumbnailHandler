using System;
using SharpShell.SharpThumbnailHandler;
using System.Runtime.InteropServices;
using SharpShell.Attributes;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CgifThumbnailHandler
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.FileExtension, ".ccg")]
    public class CCGThumbnailHandler : SharpThumbnailHandler
    {
        public CCGThumbnailHandler(){}
        
        protected override Bitmap GetThumbnailImage(uint width)
        {
            try
            {
                byte[] fileType = new byte[5];
                SelectedItemStream.Read(fileType, 0, 5);
                if (Encoding.UTF8.GetString(fileType) != "CPYCG")
                    return null;
                byte[] thumbnailLength = new byte[4];
                SelectedItemStream.Read(thumbnailLength, 0, 4);
                int length = (thumbnailLength[0] << 24) + (thumbnailLength[1] << 16)
                    + (thumbnailLength[2] << 8) + thumbnailLength[3];

                if (length == 0) return null;

                byte[] thumbnailImageData = new byte[length];
                SelectedItemStream.Read(thumbnailImageData, 0, length);
                SelectedItemStream.Close();
                return CreateThumbnailImage(thumbnailImageData);
            }
            catch (Exception ex)
            {
                LogError("An exception occured opening the CCG file.", ex);
                return null;
            }
        }

        private Bitmap CreateThumbnailImage(byte[] imagedata)
        {
            MemoryStream ms = new MemoryStream(imagedata);
            Image im = Image.FromStream(ms);
            Bitmap bt = new Bitmap(im);
            Bitmap bitmap = new Bitmap(bt.Width, bt.Height);

            using (Graphics gp = Graphics.FromImage(bitmap))
            {
                gp.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                gp.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                gp.Clear(Color.Empty);
                gp.DrawImage(bt, new Rectangle(0, 0, bitmap.Width, bitmap.Height), new Rectangle(0, 0, bt.Width, bt.Height), GraphicsUnit.Pixel);
                bt.Dispose();
                im.Dispose();
                ms.Dispose();
            }
            return bitmap;
        }
    }
}

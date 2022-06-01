using System;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace General.Librerias.CodigoUsuario
{
    public class JpgCompressor
    {
        public static byte[] Comprimir(byte[] buffer, int compresion)
        {
            byte[] rpta = null;
            Bitmap bmp = null;
            using (MemoryStream ms1 = new MemoryStream(buffer))
            {
                bmp = new Bitmap(ms1);
                ImageCodecInfo jpgEncoder = GetEncoder(ImageFormat.Jpeg);
                Encoder myEncoder = Encoder.Quality;
                EncoderParameters myEncoderParameters = new EncoderParameters(1);
                Int64 nivelCompresion = (100 - compresion);
                EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, nivelCompresion);
                myEncoderParameters.Param[0] = myEncoderParameter;
                using (MemoryStream ms = new MemoryStream())
                {
                    bmp.Save(ms, jpgEncoder, myEncoderParameters);
                    rpta = ms.ToArray();
                }
            }
            return rpta;
        }

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }
    }
}

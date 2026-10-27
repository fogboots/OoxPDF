using Lokad.OoxPdf.Imaging;
namespace Lokad.OoxPdf.Tests;
internal static class TiffImageTests
{
    // Independently encoded by Pillow/libtiff, synthetic pixels only.
    private static byte[] Raw() => Convert.FromBase64String("SUkqAAgAAAAKAAABBAABAAAAAgAAAAEBBAABAAAAAgAAAAIBAwADAAAAhgAAAAMBAwABAAAAAQAAAAYBAwABAAAAAgAAABEBBAABAAAAjAAAABUBAwABAAAAAwAAABYBBAABAAAAAgAAABcBBAABAAAADAAAABwBAwABAAAAAQAAAAAAAAAIAAgACAAKFB4oMjxGUFpkbng=");
    private static byte[] Lzw() => Convert.FromBase64String("SUkqAA4DAACAACBAMAgGCQaCweFQmGQiHQuHw2IROJRWIxeKRiLRmORuPRqQR2Qx+RSWSSeRymRAEDgKTSqUS+ZTGaTCbTObzWcTudT2cxcBA4CTyf0SjT6j0WkUulU2k0+JBUDVCnUyqVerVmq1usVyGgQOAiu2OtWSvWezWmyzUCiMFWq0Wu43O4XW5V0DCoGXa6Xe/X3AXzBRYDjEHYPA3/EYvFY2tggcBDGYnKZPLY7EAkfhLL5XMZ/PaHOgojBTO6fQajRamfAsmhbVbHWavabKHAwqBjbbvZ7za44GlsNb7ib3jb/Qg4xBzi8jm8/jy8HmkPdDndHsdfKBA4CDrd/s+DkBE7iLxefw+ntRAJH4Sej1+r4Z0JoUTfP5fn4/uMhRGBQ/D+P1AK4AqSYVQJAcFQE7QLE0FkEwZCMJoaC5QhdCkJQXDKcAwVAYQ5DcRQ0wAMleGUQxJFLxA0WwaRXEcYRUgoNl6G0ZRxGMFg4YgcRzGcgQyDplh1H8jR07QPGkHkjyDJEAg+bIfSbKkntmEBwCBKsnS48QQnOIUtzFKzABEdwiTHLsyOaEZ6iNNM4TWrQSH4JE4zVPDbBKAYlTvP05J+EwFCZP880AzoTgiJ1C0ZQ6PhQDAoUbQ1KMQFIPilSdNUchwVBMKlN0rTi4BWForVDVFABYGgsVTUVXq6Fodi1V1ayQFwhC5W1YVGrQXiSL1d2FBYYCgMFh15ZKcBiK4xWRZ71BkLwyWhZVepsGYyjNatuOOGg2DRbtrXGiAajmNVxXS0AbD0Nl1XJa6IBuQI3Xfey/hwRA4XveFkhyR45X5gStB0Sw6YHfsAh2To7YRhyfB4Ug8YfhNAB6VY9YpjSUB8WQ+Y3isuB+XI/ZBkyLCAYBAZPkLmiCY5BZZmSCiEZxCZnlsFiGapDZxlgiG4RGfXjMginGRWh5y/YjHURmk5AI54kdp+iSAJB8EhqmlM6JJ/klrWKCUAxKbBqrfCWBpLbLrbjiYChMbXhAmg2TW47M/QnBETm7bY0InhST2+X4gIALAAABAwABAAAAYAAAAAEBAwABAAAAUAAAAAIBAwADAAAAmAMAAAMBAwABAAAABQAAAAYBAwABAAAAAgAAABEBBAABAAAACAAAABUBAwABAAAAAwAAABYBAwABAAAAUAAAABcBBAABAAAABgMAABwBAwABAAAAAQAAAD0BAwABAAAAAgAAAAAAAAAIAAgACAA=");
    public static void TiffRgbAndLzwPredictorDecodeExactPixels()
    {
        using var small = TiffImage.ReadOwned(Raw());
        TestAssert.Equal(2, small.Width);
        TestAssert.True(small.Rgb.SequenceEqual(new byte[] {10,20,30,40,50,60,70,80,90,100,110,120}), "Raw RGB pixels differ.");
        using var large = TiffImage.ReadOwned(Lzw());
        TestAssert.Equal(96, large.Width);
        TestAssert.Equal(80, large.Height);
        for (int y = 0; y < 80; y++) for (int x = 0; x < 96; x++)
        {
            int offset = (y * 96 + x) * 3;
            TestAssert.Equal((byte)((x*3+y)%256), large.Rgb[offset]);
            TestAssert.Equal((byte)((y*7+x)%256), large.Rgb[offset+1]);
            TestAssert.Equal((byte)((x+y*2)%256), large.Rgb[offset+2]);
        }
    }
    public static void TiffAssociatedAlphaIsUnpremultiplied()
    {
        using var pixels = TiffImage.ReadOwned(Convert.FromBase64String("SUkqAAgAAAALAAABBAABAAAAAgAAAAEBBAABAAAAAQAAAAIBAwAEAAAAkgAAAAMBAwABAAAAAQAAAAYBAwABAAAAAgAAABEBBAABAAAAmgAAABUBAwABAAAABAAAABYBBAABAAAAAQAAABcBBAABAAAACAAAABwBAwABAAAAAQAAAFIBAwABAAAAAQAAAAAAAAAIAAgACAAIAEAgEIAAAAAA"));
        TestAssert.True(pixels.Rgb.SequenceEqual(new byte[] {127,63,31,0,0,0}), "Associated alpha must be unpremultiplied.");
        TestAssert.True(pixels.Alpha!.SequenceEqual(new byte[] {128,0}), "Transparency differs.");
    }
    public static void TiffRejectsExcessPixelsAndHostileDimensions()
    {
        byte[] bytes = Raw();
        // Independently encoded fixture has its IFD at offset 8; locate SHORT width tag.
        int count = BitConverter.ToUInt16(bytes, 8);
        for (int i = 0; i < count; i++)
        {
            int entry = 10 + 12*i;
            if (BitConverter.ToUInt16(bytes, entry) != 256) continue;
            bytes[entry+8] = 1; bytes[entry+9] = 0;
            TestAssert.Throws<InvalidDataException>(() => TiffImage.ReadOwned(bytes));
            bytes[entry+8] = 255; bytes[entry+9] = 255;
            TestAssert.Throws<InvalidDataException>(() => TiffImage.ReadOwned(bytes));
        }
    }
    public static void TiffRejectsTruncatedInputAndCancellation()
    {
        byte[] bytes = Raw();
        for (int length = 0; length < bytes.Length; length++)
            TestAssert.Throws<InvalidDataException>(() => TiffImage.ReadOwned(bytes[..length]));
        TestAssert.Throws<OperationCanceledException>(() => TiffImage.ReadOwned(bytes, new CancellationToken(true)));
        TestAssert.True(OoxImageDecoder.IsSupportedContentType("IMAGE/TIFF"), "TIFF dispatch missing.");
    }
}

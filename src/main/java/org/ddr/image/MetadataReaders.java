package org.ddr.image;

import org.ddr.image.bmp.BmpMetadataReader;
import org.ddr.image.eps.EpsMetadataReader;
import org.ddr.image.gif.GifMetadataReader;
import org.ddr.image.heif.HeifMetadataReader;
import org.ddr.image.jpeg.JpegMetadataReader;
import org.ddr.image.png.PngMetadataReader;
import org.ddr.image.tiff.TiffMetadataReader;
import org.ddr.image.webp.WebpMetadataReader;

public class MetadataReaders {
    public static final MetadataReader[] INSTANCES = {
            new JpegMetadataReader(),
            new PngMetadataReader(),
            new GifMetadataReader(),
            new WebpMetadataReader(),
            new HeifMetadataReader(),
            new BmpMetadataReader(),
            new TiffMetadataReader(),
            new EpsMetadataReader()
    };
}

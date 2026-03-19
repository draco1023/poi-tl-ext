package org.ddr.image.avif;

import org.ddr.image.heif.HeifImageReader;

import javax.imageio.spi.ImageReaderSpi;

public class AvifImageReader extends HeifImageReader {

    public AvifImageReader(ImageReaderSpi provider) {
        super(provider, new AvifMetadataReader(), "avif");
    }
}

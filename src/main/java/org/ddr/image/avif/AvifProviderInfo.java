package org.ddr.image.avif;

import com.twelvemonkeys.imageio.spi.ReaderWriterProviderInfo;

final class AvifProviderInfo extends ReaderWriterProviderInfo {
    AvifProviderInfo() {
        super(
                AvifProviderInfo.class,
                new String[]{"avif", "AVIF"}, // Names
                new String[]{"avif", "avifs"}, // Suffixes
                new String[]{"image/avif", "image/avifs"}, // Mime-types
                "org.ddr.image.avif.AvifImageReader", // Reader class name
                new String[]{"org.ddr.image.avif.AvifImageReaderSpi"},
                null,
                null,
                false, null, null, null, null,
                true, null, null, null, null
        );
    }
}

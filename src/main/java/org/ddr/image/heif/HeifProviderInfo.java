package org.ddr.image.heif;

import com.twelvemonkeys.imageio.spi.ReaderWriterProviderInfo;

final class HeifProviderInfo extends ReaderWriterProviderInfo {
    HeifProviderInfo() {
        super(
                HeifProviderInfo.class,
                new String[]{"heif", "HEIF"}, // Names
                new String[]{"heif", "heic"}, // Suffixes
                new String[]{"image/heif", "image/heic", "image/heif-sequence", "image/heic-sequence"}, // Mime-types
                "org.ddr.image.heif.HeifImageReader", // Reader class name
                new String[]{"org.ddr.image.heif.HeifImageReaderSpi"},
                null,
                null,
                false, null, null, null, null,
                true, null, null, null, null
        );
    }
}

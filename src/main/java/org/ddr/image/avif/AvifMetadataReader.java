package org.ddr.image.avif;

import com.drew.imaging.FileType;
import org.ddr.image.heif.HeifMetadataReader;

public class AvifMetadataReader extends HeifMetadataReader {
    @Override
    public boolean canRead(FileType type) {
        // FIXME metadata-extractor 一直未发版支持 AVIF 格式，会被归为 QuickTime 格式
        return type == FileType.QuickTime || type == FileType.Heif;
    }

}

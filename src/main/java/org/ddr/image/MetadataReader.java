package org.ddr.image;

import com.drew.imaging.FileType;
import com.drew.metadata.Metadata;

import java.awt.*;

public interface MetadataReader {
    boolean canRead(FileType type);

    ImageType getType(Metadata metadata);

    Dimension getDimension(Metadata metadata);
}

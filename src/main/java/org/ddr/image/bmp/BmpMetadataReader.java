package org.ddr.image.bmp;

import com.drew.imaging.FileType;
import com.drew.metadata.Directory;
import com.drew.metadata.Metadata;
import com.drew.metadata.bmp.BmpHeaderDirectory;
import org.ddr.image.ImageType;
import org.ddr.image.MetadataReader;

import java.awt.*;

public class BmpMetadataReader implements MetadataReader {
    @Override
    public boolean canRead(FileType type) {
        return type == FileType.Bmp;
    }

    @Override
    public ImageType getType(Metadata metadata) {
        return ImageType.BMP;
    }

    @Override
    public Dimension getDimension(Metadata metadata) {
        for (Directory directory : metadata.getDirectories()) {
            if (directory instanceof BmpHeaderDirectory) {
                Integer width = directory.getInteger(BmpHeaderDirectory.TAG_IMAGE_WIDTH);
                Integer height = directory.getInteger(BmpHeaderDirectory.TAG_IMAGE_HEIGHT);
                if (width != null && height != null) {
                    return new Dimension(width, height);
                }
            }
        }
        return null;
    }
}

package org.ddr.image.jpeg;

import com.drew.imaging.FileType;
import com.drew.metadata.Directory;
import com.drew.metadata.Metadata;
import com.drew.metadata.jpeg.JpegDirectory;
import org.ddr.image.ImageType;
import org.ddr.image.MetadataReader;

import java.awt.*;

public class JpegMetadataReader implements MetadataReader {
    @Override
    public boolean canRead(FileType type) {
        return type == FileType.Jpeg;
    }

    @Override
    public ImageType getType(Metadata metadata) {
        return ImageType.JPG;
    }

    @Override
    public Dimension getDimension(Metadata metadata) {
        for (Directory directory : metadata.getDirectories()) {
            if (directory instanceof JpegDirectory) {
                Integer width = directory.getInteger(JpegDirectory.TAG_IMAGE_WIDTH);
                Integer height = directory.getInteger(JpegDirectory.TAG_IMAGE_HEIGHT);
                if (width != null && height != null) {
                    return new Dimension(width, height);
                }
            }
        }
        return null;
    }
}

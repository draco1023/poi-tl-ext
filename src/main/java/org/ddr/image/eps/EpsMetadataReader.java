package org.ddr.image.eps;

import com.drew.imaging.FileType;
import com.drew.metadata.Directory;
import com.drew.metadata.Metadata;
import com.drew.metadata.eps.EpsDirectory;
import org.ddr.image.ImageType;
import org.ddr.image.MetadataReader;

import java.awt.*;

public class EpsMetadataReader implements MetadataReader{
    @Override
    public boolean canRead(FileType type) {
        return type == FileType.Eps;
    }

    @Override
    public ImageType getType(Metadata metadata) {
        return ImageType.EPS;
    }

    @Override
    public Dimension getDimension(Metadata metadata) {
        for (Directory directory : metadata.getDirectories()) {
            if (directory instanceof EpsDirectory) {
                Integer width = directory.getInteger(EpsDirectory.TAG_IMAGE_WIDTH);
                Integer height = directory.getInteger(EpsDirectory.TAG_IMAGE_HEIGHT);
                if (width != null && height != null) {
                    return new Dimension(width, height);
                }
            }
        }
        return null;
    }
}

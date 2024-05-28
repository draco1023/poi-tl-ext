package org.ddr.image.webp;

import com.drew.imaging.FileType;
import com.drew.metadata.Directory;
import com.drew.metadata.Metadata;
import com.drew.metadata.webp.WebpDirectory;
import org.ddr.image.ImageType;
import org.ddr.image.MetadataReader;

import java.awt.*;

public class WebpMetadataReader implements MetadataReader{
    @Override
    public boolean canRead(FileType type) {
        return type == FileType.WebP;
    }

    @Override
    public ImageType getType(Metadata metadata) {
        for (Directory directory : metadata.getDirectories()) {
            if (directory instanceof WebpDirectory) {
                Boolean hasAlpha = directory.getBooleanObject(WebpDirectory.TAG_HAS_ALPHA);
                if (Boolean.TRUE.equals(hasAlpha)) {
                    return ImageType.PNG;
                }
            }
        }
        return ImageType.JPG;
    }

    @Override
    public Dimension getDimension(Metadata metadata) {
        for (Directory directory : metadata.getDirectories()) {
            if (directory instanceof WebpDirectory) {
                Integer width = directory.getInteger(WebpDirectory.TAG_IMAGE_WIDTH);
                Integer height = directory.getInteger(WebpDirectory.TAG_IMAGE_HEIGHT);
                if (width != null && height != null) {
                    return new Dimension(width, height);
                }
            }
        }
        return null;
    }
}

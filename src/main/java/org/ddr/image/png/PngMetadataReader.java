package org.ddr.image.png;

import com.drew.imaging.FileType;
import com.drew.imaging.png.PngChunkType;
import com.drew.metadata.Directory;
import com.drew.metadata.Metadata;
import com.drew.metadata.png.PngDirectory;
import org.ddr.image.ImageType;
import org.ddr.image.MetadataReader;

import java.awt.*;

public class PngMetadataReader implements MetadataReader {
    @Override
    public boolean canRead(FileType type) {
        return type == FileType.Png;
    }

    @Override
    public ImageType getType(Metadata metadata) {
        return ImageType.PNG;
    }

    @Override
    public Dimension getDimension(Metadata metadata) {
        for (Directory directory : metadata.getDirectories()) {
            if (directory instanceof PngDirectory) {
                if (((PngDirectory) directory).getPngChunkType() == PngChunkType.IHDR) {
                    Integer width = directory.getInteger(PngDirectory.TAG_IMAGE_WIDTH);
                    Integer height = directory.getInteger(PngDirectory.TAG_IMAGE_HEIGHT);
                    if (width != null && height != null) {
                        return new Dimension(width, height);
                    }
                }
            }
        }
        return null;
    }
}

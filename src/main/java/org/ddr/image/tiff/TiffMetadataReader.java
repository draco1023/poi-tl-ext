package org.ddr.image.tiff;

import com.drew.imaging.FileType;
import com.drew.metadata.Directory;
import com.drew.metadata.Metadata;
import com.drew.metadata.exif.ExifIFD0Directory;
import org.ddr.image.ImageType;
import org.ddr.image.MetadataReader;

import java.awt.*;
import java.util.EnumSet;

public class TiffMetadataReader implements MetadataReader {
    private static final EnumSet<FileType> TIFF_TYPES = EnumSet.of(FileType.Tiff, FileType.Arw, FileType.Cr2, FileType.Nef, FileType.Orf, FileType.Rw2);

    @Override
    public boolean canRead(FileType type) {
        return TIFF_TYPES.contains(type);
    }

    @Override
    public ImageType getType(Metadata metadata) {
        return ImageType.TIFF;
    }

    @Override
    public Dimension getDimension(Metadata metadata) {
        for (Directory directory : metadata.getDirectories()) {
            if (directory instanceof ExifIFD0Directory) {
                Integer width = directory.getInteger(ExifIFD0Directory.TAG_IMAGE_WIDTH);
                Integer height = directory.getInteger(ExifIFD0Directory.TAG_IMAGE_HEIGHT);
                if (width != null && height != null) {
                    return new Dimension(width, height);
                }
            }
        }
        return null;
    }
}

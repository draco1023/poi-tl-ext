package org.ddr.image.heif;

import com.drew.imaging.FileType;
import com.drew.metadata.Directory;
import com.drew.metadata.Metadata;
import com.drew.metadata.exif.ExifDescriptorBase;
import com.drew.metadata.exif.ExifIFD0Directory;
import com.drew.metadata.exif.ExifSubIFDDirectory;
import com.drew.metadata.heif.HeifDescriptor;
import com.drew.metadata.heif.HeifDirectory;
import org.ddr.image.ImageType;
import org.ddr.image.MetadataReader;

import java.awt.*;

public class HeifMetadataReader implements MetadataReader{
    @Override
    public boolean canRead(FileType type) {
        return type == FileType.Heif;
    }

    @Override
    public ImageType getType(Metadata metadata) {
        // FIXME read icc profile
        return ImageType.JPG;
    }

    /**
     * @see HeifDescriptor#getRotationDescription()
     * @see ExifDescriptorBase#getOrientationDescription()
     */
    @Override
    public Dimension getDimension(Metadata metadata) {
        Integer width = null;
        Integer height = null;
        Boolean rotated = null;
        for (Directory directory : metadata.getDirectories()) {
            if (directory instanceof HeifDirectory) {
                Integer r = directory.getInteger(HeifDirectory.TAG_IMAGE_ROTATION);
                if (r != null) {
                    rotated = r % 2 == 1;
                }
            } else if (directory instanceof ExifIFD0Directory) {
                if (rotated == null) {
                    Integer r = directory.getInteger(ExifIFD0Directory.TAG_ORIENTATION);
                    if (r != null) {
                        rotated = r > 4;
                    }
                }
            } else if (directory instanceof ExifSubIFDDirectory) {
                width = directory.getInteger(ExifSubIFDDirectory.TAG_EXIF_IMAGE_WIDTH);
                height = directory.getInteger(ExifSubIFDDirectory.TAG_EXIF_IMAGE_HEIGHT);
            }
        }
        if (width != null && height != null) {
            if (Boolean.TRUE.equals(rotated)) {
                //noinspection SuspiciousNameCombination
                return new Dimension(height, width);
            }
            return new Dimension(width, height);
        }
        return null;
    }
}

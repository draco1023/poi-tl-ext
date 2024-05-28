package org.ddr.image;

import java.awt.*;
import java.io.ByteArrayInputStream;

public class ImageInfo {
    private ByteArrayInputStream stream;
    private ImageType type;
    private Dimension dimension;

    public ImageInfo(ByteArrayInputStream stream) {
        this.stream = stream;
    }

    public ImageInfo(ByteArrayInputStream stream, ImageType type, Dimension dimension) {
        this.stream = stream;
        this.type = type;
        this.dimension = dimension;
    }

    public ByteArrayInputStream getStream() {
        return stream;
    }

    public void setStream(ByteArrayInputStream stream) {
        this.stream = stream;
    }

    public ImageType getType() {
        return type;
    }

    public void setType(ImageType type) {
        this.type = type;
    }

    public Dimension getDimension() {
        return dimension;
    }

    public void setDimension(Dimension dimension) {
        this.dimension = dimension;
    }

    public int getWidth() {
        return dimension == null ? 0 : dimension.width;
    }

    public int getHeight() {
        return dimension == null ? 0 : dimension.height;
    }

    public int getRawType() {
        return type == null ? -1 : type.getType();
    }
}

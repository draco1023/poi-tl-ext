package org.ddr.image;

import org.apache.poi.xwpf.usermodel.Document;

public enum ImageType {
    EMF(Document.PICTURE_TYPE_EMF),
    WMF(Document.PICTURE_TYPE_WMF),
    PICT(Document.PICTURE_TYPE_PICT),
    JPEG(Document.PICTURE_TYPE_JPEG),
    JPG(Document.PICTURE_TYPE_JPEG),
    PNG(Document.PICTURE_TYPE_PNG),
    DIB(Document.PICTURE_TYPE_DIB),
    GIF(Document.PICTURE_TYPE_GIF),
    TIF(Document.PICTURE_TYPE_TIFF),
    TIFF(Document.PICTURE_TYPE_TIFF),
    EPS(Document.PICTURE_TYPE_EPS),
    BMP(Document.PICTURE_TYPE_BMP),
    WPG(Document.PICTURE_TYPE_WPG);

    private final int type;

    ImageType(int type) {
        this.type = type;
    }

    public String getExtension() {
        return name().toLowerCase();
    }

    public int getType() {
        return type;
    }
}

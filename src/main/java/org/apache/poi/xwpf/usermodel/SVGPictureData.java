package org.apache.poi.xwpf.usermodel;

import org.apache.poi.openxml4j.opc.PackagePart;

/**
 * SVGPictureData
 *
 * @author Draco
 * @since 2022-04-12
 */
public class SVGPictureData extends XWPFPictureData {

    public static final int PICTURE_TYPE_SVG = 1;

    public SVGPictureData() {
    }

    public SVGPictureData(PackagePart part) {
        super(part);
    }

    public static void initRelation() {
        RELATIONS[PICTURE_TYPE_SVG] = SVGRelation.INSTANCE;
    }
}

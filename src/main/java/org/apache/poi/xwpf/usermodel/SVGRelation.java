package org.apache.poi.xwpf.usermodel;

import org.apache.poi.ooxml.POIXMLRelation;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.sl.usermodel.PictureData;

import javax.xml.namespace.QName;

/**
 * SVGRelation
 *
 * @author Draco
 * @since 2022-04-12
 */
public class SVGRelation extends POIXMLRelation {
    public static final SVGRelation INSTANCE = new SVGRelation();
    /**
     * @see org.apache.poi.xslf.usermodel.XSLFPictureShape#MS_SVG_NS
     */
    public static final String MS_SVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
    /**
     * @see org.apache.poi.xslf.usermodel.XSLFPictureShape#SVG_URI
     */
    public static final String SVG_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
    /**
     * @see org.apache.poi.xslf.usermodel.XSLFPictureShape#EMBED_TAG
     */
    public static final QName EMBED_TAG = new QName(PackageRelationshipTypes.CORE_PROPERTIES_ECMA376_NS,
            "embed", "r");

    public static final String SVG_BLIP = "svgBlip";

    public static final String SVG_PREFIX = "asvg";

    public static final QName SVG_QNAME = new QName(MS_SVG_NS, SVG_BLIP, SVG_PREFIX);

    /**
     * @see XWPFRelation#IMAGE_PNG
     */
    private SVGRelation() {
        super(PictureData.PictureType.SVG.contentType,
                PackageRelationshipTypes.IMAGE_PART,
                "/word/media/image#.svg",
                SVGPictureData::new, SVGPictureData::new, null);
    }
}

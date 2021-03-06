/*
 * Copyright 2016 - 2021 Draco, https://github.com/draco1023
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.ddr.poi.html;

import com.deepoove.poi.render.RenderContext;
import com.steadystate.css.dom.CSSStyleDeclarationImpl;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BodyType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.SVGPictureData;
import org.apache.poi.xwpf.usermodel.SVGRelation;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.ddr.poi.html.util.CSSLength;
import org.ddr.poi.html.util.CSSLengthUnit;
import org.ddr.poi.html.util.CSSStyleUtils;
import org.ddr.poi.html.util.Colors;
import org.ddr.poi.html.util.InlineStyle;
import org.ddr.poi.html.util.NamedFontSize;
import org.ddr.poi.html.util.NumberingContext;
import org.ddr.poi.html.util.RenderUtils;
import org.ddr.poi.html.util.WhiteSpaceRule;
import org.ddr.poi.util.XmlUtils;
import org.jsoup.internal.StringUtil;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlip;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObjectFrameLocking;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualGraphicFrameProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTOfficeArtExtension;
import org.openxmlformats.schemas.drawingml.x2006.main.CTOfficeArtExtensionList;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STThemeColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalAlignRun;

import javax.xml.namespace.QName;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.LinkedList;

/**
 * HTML????????????????????????
 *
 * @author Draco
 * @since 2021-02-08
 */
public class HtmlRenderContext extends RenderContext<String> {
    private static final QName R_QNAME = new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "r");
    private static final QName HYPERLINK_QNAME = new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "hyperlink");

    /**
     * ???????????? ?????? 12pt 16px
     */
    private static final CSSLength DEFAULT_FONT_SIZE = new CSSLength(12, CSSLengthUnit.PT);
    /**
     * ?????????????????????
     */
    private static final String DEFAULT_HYPERLINK_COLOR = "0563C1";

    /**
     * ?????????IBodyElement????????????????????????HTML?????????IBodyElement???????????????
     */
    private LinkedList<IBodyElement> closestBodyStack = new LinkedList<>();

    /**
     * ????????????????????????????????????/?????????????????????????????????HTML???????????????????????????
     */
    private LinkedList<IBody> ancestors = new LinkedList<>();

    /**
     * ???????????????????????????????????????????????????
     */
    private LinkedList<InlineStyle> inlineStyles = new LinkedList<>();

    /**
     * ????????????????????????????????????????????????????????????
     */
    private LinkedList<Integer> fontSizesInHalfPoints = new LinkedList<>();

    /**
     * ??????????????????????????????????????????
     */
    private final NumberingContext numberingContext;
    /**
     * ????????????
     */
    private final CSSLength defaultFontSize;
    /**
     * ????????????
     */
    private final CSSLength pageWidth;
    /**
     * ????????????
     */
    private final CSSLength pageHeight;
    /**
     * ??????????????????
     */
    private final CSSLength marginTop;
    /**
     * ??????????????????
     */
    private final CSSLength marginRight;
    /**
     * ??????????????????
     */
    private final CSSLength marginBottom;
    /**
     * ??????????????????
     */
    private final CSSLength marginLeft;
    /**
     * ??????????????????
     */
    private final int availablePageWidth;
    /**
     * ??????????????????
     */
    private final int availablePageHeight;
    /**
     * ??????????????????????????????ID
     */
    private String placeholderStyleId;

    /**
     * ??????Run???????????????????????????
     */
    private XWPFRun currentRun;

    /**
     * ?????????????????????????????????????????????????????????
     */
    private String globalFont;
    /**
     * ?????????????????????????????????????????????????????????
     */
    private BigInteger globalFontSize;

    /**
     * ???????????????????????????
     */
    private int blockLevel;

    /**
     * ????????????
     *
     * @param context ?????????????????????
     */
    public HtmlRenderContext(RenderContext<String> context) {
        super(context.getEleTemplate(), context.getData(), context.getTemplate());
        numberingContext = new NumberingContext(getXWPFDocument());

        CTSectPr sectPr = getXWPFDocument().getDocument().getBody().getSectPr();
        CTPageSz pgSz = sectPr.getPgSz();
        // ?????????????????????twip
        int w = pgSz.getW().intValue();
        pageWidth = new CSSLength(w, CSSLengthUnit.TWIP);
        int h = pgSz.getH().intValue();
        pageHeight = new CSSLength(h, CSSLengthUnit.TWIP);

        CTPageMar pgMar = sectPr.getPgMar();
        int top = pgMar.getTop().intValue();
        marginTop = new CSSLength(top, CSSLengthUnit.TWIP);
        int right = pgMar.getRight().intValue();
        marginRight = new CSSLength(right, CSSLengthUnit.TWIP);
        int bottom = pgMar.getBottom().intValue();
        marginBottom = new CSSLength(bottom, CSSLengthUnit.TWIP);
        int left = pgMar.getLeft().intValue();
        marginLeft = new CSSLength(left, CSSLengthUnit.TWIP);

        availablePageWidth = new CSSLength(w - left - right, CSSLengthUnit.TWIP).toEMU();
        availablePageHeight = new CSSLength(h - top - bottom, CSSLengthUnit.TWIP).toEMU();

        int fontSize = getXWPFDocument().getStyles().getDefaultRunStyle().getFontSize();
        defaultFontSize = fontSize > 0 ? new CSSLength(fontSize, CSSLengthUnit.PT) : DEFAULT_FONT_SIZE;

        extractPlaceholderStyle();
    }

    /**
     * ????????????????????????????????????
     */
    private void extractPlaceholderStyle() {
        XWPFRun run = getRun();
        IRunBody runParent = run.getParent();
        if (runParent instanceof XWPFParagraph) {
            XWPFParagraph paragraph = (XWPFParagraph) runParent;
            String styleId = paragraph.getStyleID();
            boolean existsRPr = run.getCTR().isSetRPr();

            if (styleId == null && !existsRPr) {
                return;
            } else if (styleId != null && !existsRPr) {
                placeholderStyleId = styleId;
                return;
            }

            XWPFStyles styles = getXWPFDocument().getStyles();
            CTStyle newCTStyle = CTStyle.Factory.newInstance();
            newCTStyle.setCustomStyle(STOnOff.TRUE);
            newCTStyle.setType(STStyleType.PARAGRAPH);
            newCTStyle.addNewHidden();
            newCTStyle.setRPr(run.getCTR().getRPr());
            XmlUtils.removeNamespaces(newCTStyle.getRPr());

            String newStyleId = styleId + getXWPFDocument().getStyles().getNumberOfStyles();
            newCTStyle.setStyleId(newStyleId);
            newCTStyle.addNewName().setVal(newStyleId);
            placeholderStyleId = newStyleId;

            if (styleId != null) {
                newCTStyle.addNewBasedOn().setVal(styleId);
            }

            XWPFStyle newStyle = new XWPFStyle(newCTStyle, styles);
            styles.addStyle(newStyle);

            paragraph.setStyle(newStyleId);
        }
    }

    @Override
    public IBody getContainer() {
        IBody container = ancestors.peek();
        return container == null ? super.getContainer() : container;
    }

    public boolean containerChanged() {
        return !ancestors.isEmpty();
    }

    /**
     * ???????????????
     *
     * @param body ?????????
     */
    public void pushContainer(IBody body) {
        ancestors.push(body);
    }

    /**
     * ???????????????
     */
    public void popContainer() {
        ancestors.pop();
    }

    /**
     * ??????????????????????????????????????????????????????
     *
     * @return ???????????????????????????
     */
    public IBodyElement getClosestBody() {
        IBodyElement body = closestBodyStack.peek();
        return body == null ? (IBodyElement) getRun().getParent() : body;
    }

    /**
     * ???????????????????????????????????????
     *
     * @param body ?????????????????????????????????
     */
    public void pushClosestBody(IBodyElement body) {
        closestBodyStack.push(body);
    }

    /**
     * ???????????????????????????????????????
     */
    public void popClosestBody() {
        closestBodyStack.pop();
    }

    /**
     * ??????????????????????????????????????????????????????
     *
     * @param body ?????????????????????????????????
     */
    public void replaceClosestBody(IBodyElement body) {
        if (!closestBodyStack.isEmpty()) {
            closestBodyStack.pop();
        }
        closestBodyStack.push(body);
    }

    /**
     * ???????????????????????????????????????????????????????????????????????????????????????????????????????????????
     *
     * @return ???????????????
     */
    public XWPFParagraph getClosestParagraph() {
        IBodyElement body = closestBodyStack.peek();
        if (body == null) {
            return (XWPFParagraph) getRun().getParent();
        }
        switch (body.getElementType()) {
            case PARAGRAPH:
                return (XWPFParagraph) body;
            case TABLE:
                XmlCursor xmlCursor = ((XWPFTable) body).getCTTbl().newCursor();
                xmlCursor.toEndToken();
                xmlCursor.toNextToken();
                XWPFParagraph paragraph = newParagraph(null, xmlCursor);
                replaceClosestBody(paragraph);
                xmlCursor.dispose();
                return paragraph;
        }
        throw new IllegalStateException("Impossible");
    }

    /**
     * ?????????????????????
     *
     * @param uri ????????????
     */
    public void startHyperlink(String uri) {
        if (isBlocked()) {
            currentRun = getClosestParagraph().createHyperlinkRun(uri);
        } else {
            // ?????????????????????????????????
            String rId = getRun().getParent().getPart().getPackagePart()
                    .addExternalRelationship(uri, XWPFRelation.HYPERLINK.getRelation()).getId();
            XmlCursor xmlCursor = getRun().getCTR().newCursor();
            xmlCursor.insertElement(HYPERLINK_QNAME);
            xmlCursor.toPrevSibling();
            CTHyperlink ctHyperlink = (CTHyperlink) xmlCursor.getObject();
            ctHyperlink.setId(rId);
            ctHyperlink.addNewR();
            currentRun = new XWPFHyperlinkRun(ctHyperlink, ctHyperlink.getRArray(0), getRun().getParent());
        }
    }

    /**
     * ?????????????????????
     */
    public void endHyperlink() {
        currentRun = null;
    }

    /**
     * ????????????
     *
     * @param container ??????
     * @param cursor xml??????
     * @return ??????
     */
    public XWPFParagraph newParagraph(IBody container, XmlCursor cursor) {
        if (container == null) {
            container = getContainer();
        }
        XWPFParagraph xwpfParagraph = container.insertNewParagraph(cursor);
        if (placeholderStyleId != null) {
            xwpfParagraph.setStyle(placeholderStyleId);
        }
        return xwpfParagraph;
    }

    /**
     * ??????CTR
     *
     * @return CTR
     */
    public CTR newRun() {
        // ????????????????????????????????????????????????????????????run
        if (currentRun instanceof XWPFHyperlinkRun) {
            XmlCursor xmlCursor = currentRun.getCTR().newCursor();
            CTR ctr;
            if (xmlCursor.toFirstChild()) {
                xmlCursor.dispose();
                ctr = ((XWPFHyperlinkRun) currentRun).getCTHyperlink().addNewR();
            } else {
                // run???????????????????????????
                ctr = currentRun.getCTR();
            }
            // ??????????????????
            initHyperlinkStyle(ctr);

            return ctr;
        }
        // ???????????????????????????????????????????????????run
        if (isBlocked()) {
            currentRun = getClosestParagraph().createRun();
        } else {
            // ????????????????????????run
            XmlCursor xmlCursor = getRun().getCTR().newCursor();
            xmlCursor.insertElement(R_QNAME);
            xmlCursor.toPrevSibling();
            CTR ctr = (CTR) xmlCursor.getObject();
            currentRun = new XWPFRun(ctr, getRun().getParent());
        }
        return currentRun.getCTR();
    }

    /**
     * ????????????????????????
     *
     * @param ctr CTR
     */
    private void initHyperlinkStyle(CTR ctr) {
        CTRPr rPr = RenderUtils.getRPr(ctr);
        CTColor ctColor = rPr.addNewColor();
        ctColor.setVal(DEFAULT_HYPERLINK_COLOR);
        ctColor.setThemeColor(STThemeColor.HYPERLINK);

        rPr.addNewU().setVal(STUnderline.SINGLE);
    }

    /**
     * ??????????????????????????????????????????????????????????????????????????????
     *
     * @return ???????????????
     */
    public XWPFTable getClosestTable() {
        for (IBodyElement body : closestBodyStack) {
            if (body instanceof XWPFTable) {
                return ((XWPFTable) body);
            }
        }
        throw new IllegalStateException("No table in stack");
    }

    /**
     * ??????????????????
     *
     * @param inlineStyle ????????????
     * @param block ?????????????????????
     */
    public void pushInlineStyle(CSSStyleDeclarationImpl inlineStyle, boolean block) {
        String newFontSize = inlineStyle.getFontSize();
        // ??????????????????????????????
        int fontSize = Integer.MIN_VALUE;
        if (StringUtils.isNotBlank(newFontSize)) {
            NamedFontSize namedFontSize = NamedFontSize.of(newFontSize);
            if (namedFontSize != null) {
                // ?????????????????????
                fontSize = namedFontSize.getSize().toHalfPoints();
            } else if (HtmlConstants.SMALLER.equalsIgnoreCase(newFontSize)) {
                // ???????????????
                int inheritedFontSize = getInheritedFontSizeInHalfPoints();
                fontSize = RenderUtils.smallerFontSizeInHalfPoints(inheritedFontSize);
            } else if (HtmlConstants.LARGER.equalsIgnoreCase(newFontSize)) {
                // ???????????????
                int inheritedFontSize = getInheritedFontSizeInHalfPoints();
                fontSize = RenderUtils.largerFontSizeInHalfPoints(inheritedFontSize);
            } else {
                CSSLength cssLength = CSSLength.of(newFontSize);
                if (cssLength.isValid()) {
                    if (cssLength.getUnit() == CSSLengthUnit.PERCENT) {
                        fontSize = (int) Math.rint(getInheritedFontSizeInHalfPoints()
                                * cssLength.getValue() * cssLength.getUnit().absoluteFactor());
                    } else {
                        int emu = lengthToEMU(cssLength);
                        fontSize = emu * 2 / Units.EMU_PER_POINT;
                    }
                }
            }
        }
        fontSizesInHalfPoints.push(fontSize);
        inlineStyles.push(new InlineStyle(inlineStyle, block));
    }

    /**
     * ??????????????????
     */
    public void popInlineStyle() {
        fontSizesInHalfPoints.pop();
        inlineStyles.pop();
    }

    /**
     * ?????????????????????????????????HTML?????????????????????????????????????????????????????????????????????????????????????????????????????????????????????
     *
     * @return ???????????????????????????
     */
    public CSSStyleDeclarationImpl currentElementStyle() {
        InlineStyle inlineStyle = inlineStyles.peek();
        return inlineStyle == null ? CSSStyleUtils.EMPTY_STYLE : inlineStyle.getDeclaration();
    }

    /**
     * ???????????????????????????????????????
     *
     * @param property ????????????
     * @return ??????????????????????????????????????????
     */
    public String getPropertyValue(String property) {
        return getPropertyValue(property, false);
    }

    /**
     * ???????????????????????????????????????
     *
     * @param property ????????????
     * @param inlineOnly ?????????????????????????????????
     * @return ??????????????????????????????????????????
     */
    public String getPropertyValue(String property, boolean inlineOnly) {
        return getPropertyValue(property, false, inlineOnly);
    }

    /**
     * ???????????????
     *
     * @param property ????????????
     * @param caseSensitive ?????????????????????????????????????????????????????????????????????????????????
     * @param inlineOnly ?????????????????????????????????
     * @return ??????????????????????????????????????????
     */
    public String getPropertyValue(String property, boolean caseSensitive, boolean inlineOnly) {
        for (InlineStyle inlineStyle : inlineStyles) {
            if (inlineOnly && inlineStyle.isBlock()) {
                break;
            }
            String propertyValue = inlineStyle.getDeclaration().getPropertyValue(property);
            if (StringUtils.isNotBlank(propertyValue)) {
                return caseSensitive ? propertyValue : propertyValue.toLowerCase();
            }
        }
        return "";
    }

    /**
     * @return Word????????????????????????
     */
    public CSSLength getDefaultFontSize() {
        return defaultFontSize;
    }

    /**
     * @return ????????????????????????????????????????????????????????????
     */
    public int getInheritedFontSizeInHalfPoints() {
        for (Integer fontSize : fontSizesInHalfPoints) {
            if (fontSize > 0) {
                return fontSize;
            }
        }
        return defaultFontSize.toHalfPoints();
    }

    /**
     * @return ??????????????????????????????EMU?????????
     */
    public int getAvailableWidthInEMU() {
        IBody container = getContainer();
        if (container.getPartType() == BodyType.DOCUMENT) {
            return availablePageWidth;
        } else {
            return RenderUtils.getAvailableWidthInEMU(container);
        }
    }

    /**
     * ??????????????????????????????EMU?????????
     *
     * @param length ????????????
     * @param maxLength ??????????????????
     * @param naturalEMU ????????????
     * @param parentEMU ???????????????
     * @return ???EMU?????????????????????
     */
    public int computeLengthInEMU(String length, String maxLength, int naturalEMU, int parentEMU) {
        int emu = naturalEMU;

        if (length.length() > 0) {
            CSSLength cssLength = CSSLength.of(length);
            if (cssLength.isValid()) {
                emu = computeLengthInEMU(cssLength, naturalEMU, parentEMU);
            }
        }

        if (maxLength.length() > 0) {
            CSSLength cssLength = CSSLength.of(maxLength);
            if (cssLength.isValid()) {
                int maxEMU = computeLengthInEMU(cssLength, naturalEMU, parentEMU);
                emu = Math.min(maxEMU, emu);
            }
        }

        return Math.min(emu, parentEMU);
    }

    /**
     * ??????????????????????????????EMU?????????
     *
     * @param cssLength ????????????
     * @param naturalEMU ????????????
     * @param parentEMU ???????????????
     * @return ???EMU?????????????????????
     */
    public int computeLengthInEMU(CSSLength cssLength, int naturalEMU, int parentEMU) {
        int length;
        if (cssLength.getUnit() == CSSLengthUnit.PERCENT) {
            if (parentEMU != Integer.MAX_VALUE) {
                length = (int) (parentEMU * cssLength.getValue() * cssLength.getUnit().absoluteFactor());
            } else {
                length = naturalEMU;
            }
        } else {
            length = lengthToEMU(cssLength);
        }
        return length;
    }

    /**
     * ????????????
     *
     * @param text ??????
     */
    public void renderText(String text) {
        String whiteSpace = getPropertyValue(HtmlConstants.CSS_WHITE_SPACE);
        WhiteSpaceRule rule = WhiteSpaceRule.of(whiteSpace);
        CTR ctr = newRun();

        StringBuilder sb = StringUtil.borrowBuilder();
        boolean mergeWhitespace = false;
        boolean reachedNonWhite = false;

        // https://en.wikipedia.org/wiki/List_of_XML_and_HTML_character_entity_references
        int len = text.length();
        int c;

        if (!rule.isKeepTrailingSpace()) {
            boolean reachedLastNonWhite = false;
            for (int i = len - 1; i >= 0; i -= Character.charCount(c)) {
                c = text.codePointAt(i);
                switch (c) {
                    case ' ':
                    case '\r':
                    case '\n':
                    case '\t':
                    case 173: // soft hyphen
                    case 8203: // zero width space
                    case 8204: // zero width non-joiner
                    case 8205: // zero width joiner
                    case 8206: // lrm
                    case 8207: // rlm
                    case 8288: // word joiner
                    case 8289: // apply function
                    case 8290: // invisible times
                    case 8291: // invisible separator
                        len = i;
                        break;
                    default:
                        if (Character.getType(c) == 16) {
                            len = i;
                        } else {
                            reachedLastNonWhite = true;
                        }
                        break;
                }
                if (reachedLastNonWhite) {
                    break;
                }
            }
        }

        for (int i = 0; i < len; i += Character.charCount(c)) {
            c = text.codePointAt(i);
            switch (c) {
                case '\r':
                    if (i + 1 < len && text.codePointAt(i + 1) == '\n') {
                        continue;
                    }
                    if (rule.isKeepLineBreak()) {
                        addText(ctr, sb);
                        ctr.addNewCr();
                    } else {
                        mergeWhitespace = true;
                    }
                    break;
                case '\n':
                    if (rule.isKeepLineBreak()) {
                        addText(ctr, sb);
                        ctr.addNewBr();
                    } else {
                        mergeWhitespace = true;
                    }
                    break;
                case ' ':
                    if (reachedNonWhite || rule.isKeepSpaceAndTab()) {
                        sb.appendCodePoint(c);
                    } else {
                        mergeWhitespace = true;
                    }
                    break;
                case '\t':
                    if (reachedNonWhite || rule.isKeepSpaceAndTab()) {
                        addText(ctr, sb);
                        ctr.addNewTab();
                    } else {
                        mergeWhitespace = true;
                    }
                    break;
                case 160: // nbsp
                case 8194: // ensp
                case 8195: // emsp
                case 8196: // emsp13
                case 8197: // emsp14
                case 8199: // numsp
                case 8200: // puncsp
                case 8201: // thinsp
                case 8202: // hairsp
                case 8287: // medium space
                    sb.append(' ');
                    mergeWhitespace = false;
                    break;
                case 173: // soft hyphen
                case 8203: // zero width space
                case 8204: // zero width non-joiner
                case 8205: // zero width joiner
                case 8206: // lrm
                case 8207: // rlm
                case 8288: // word joiner
                case 8289: // apply function
                case 8290: // invisible times
                case 8291: // invisible separator
                    continue;
                default:
                    if (Character.getType(c) == 16) {
                        continue;
                    }
                    if (mergeWhitespace) {
                        if (reachedNonWhite) {
                            sb.append(' ');
                        }
                        mergeWhitespace = false;
                    }
                    sb.appendCodePoint(c);
                    reachedNonWhite = true;
                    break;
            }
        }

        addText(ctr, sb);
        StringUtil.releaseBuilder(sb);

        // ????????????
        applyTextStyle(ctr);

        if (!(currentRun instanceof XWPFHyperlinkRun)) {
            currentRun = null;
        }
    }

    private void addText(CTR ctr, StringBuilder sb) {
        if (sb.length() > 0) {
            CTText ctText = ctr.addNewT();
            String text = sb.toString();
            ctText.setStringValue(text);
            if (text.charAt(0) == ' ' || text.charAt(sb.length() - 1) == ' ') {
                ctText.setSpace(SpaceAttribute.Space.PRESERVE);
            }
            sb.delete(0, sb.length());
        }
    }

    /**
     * ??????????????????
     *
     * @param ctr CTR
     */
    private void applyTextStyle(CTR ctr) {
        CTRPr rPr = RenderUtils.getRPr(ctr);

        // ?????????????????????????????????????????????????????????
        String fontFamily = StringUtils.isBlank(globalFont) ? getPropertyValue(HtmlConstants.CSS_FONT_FAMILY) : globalFont;
        if (StringUtils.isNotBlank(fontFamily)) {
            CTFonts ctFonts = rPr.addNewRFonts();
            // ASCII
            ctFonts.setAscii(fontFamily);
            // High ANSI
            ctFonts.setHAnsi(fontFamily);
            // Complex Script
            ctFonts.setCs(fontFamily);
            // East Asian
            ctFonts.setEastAsia(fontFamily);
        }

        // ??????
        if (globalFontSize == null) {
            String fontSize = getPropertyValue(HtmlConstants.CSS_FONT_SIZE);
            if (StringUtils.isNotBlank(fontSize)) {
                int sz = getInheritedFontSizeInHalfPoints();
                rPr.addNewSz().setVal(BigInteger.valueOf(sz));
            }
        } else {
            // ????????????????????????????????????????????????
            rPr.addNewSz().setVal(globalFontSize);
        }

        // ??????
        String fontWeight = getPropertyValue(HtmlConstants.CSS_FONT_WEIGHT);
        if (fontWeight.contains(HtmlConstants.BOLD)) {
            rPr.addNewB();
        } else if (NumberUtils.isParsable(fontWeight) && Float.parseFloat(fontWeight) > 500) {
            rPr.addNewB();
        }

        // ??????
        String fontStyle = getPropertyValue(HtmlConstants.CSS_FONT_STYLE);
        if (HtmlConstants.ITALIC.equals(fontStyle) || HtmlConstants.OBLIQUE.equals(fontStyle)) {
            rPr.addNewI();
        }

        // FIXME cssparser???????????????????????????????????????rgb?????????????????????????????????
        // ??????
        String color = getPropertyValue(HtmlConstants.CSS_COLOR);
        if (StringUtils.isNotBlank(color)) {
            String hex = Colors.fromStyle(color);
            RenderUtils.getColor(rPr).setVal(hex);
        }

        String caps = getPropertyValue(HtmlConstants.CSS_FONT_VARIANT_CAPS);
        if (HtmlConstants.SMALL_CAPS.equals(caps)) {
            rPr.addNewSmallCaps();
        }

        // ?????????/?????????
        String textDecoration = getPropertyValue(HtmlConstants.CSS_TEXT_DECORATION_LINE);
        if (StringUtils.contains(textDecoration, HtmlConstants.LINE_THROUGH)) {
            rPr.addNewStrike();
        }
        if (StringUtils.contains(textDecoration, HtmlConstants.UNDERLINE)) {
            CTUnderline ctUnderline = RenderUtils.getUnderline(rPr);
            String textDecorationStyle = getPropertyValue(HtmlConstants.CSS_TEXT_DECORATION_STYLE);
            ctUnderline.setVal(RenderUtils.underline(textDecorationStyle));
            String textDecorationColor = getPropertyValue(HtmlConstants.CSS_TEXT_DECORATION_COLOR);
            if (StringUtils.isNotBlank(textDecorationColor)) {
                String hex = Colors.fromStyle(textDecorationColor);
                ctUnderline.setColor(hex);
            }
        }

        // ?????????
        String verticalAlign = getPropertyValue(HtmlConstants.CSS_VERTICAL_ALIGN);
        if (HtmlConstants.SUPER.equals(verticalAlign)) {
            rPr.addNewVertAlign().setVal(STVerticalAlignRun.SUPERSCRIPT);
        } else if (HtmlConstants.SUB.equals(verticalAlign)) {
            rPr.addNewVertAlign().setVal(STVerticalAlignRun.SUBSCRIPT);
        }

        // FIXME ???????????????????????????????????????????????????????????????????????????

        // ?????????
        String visibility = getPropertyValue(HtmlConstants.CSS_VISIBILITY);
        if (HtmlConstants.HIDDEN.equals(visibility) || HtmlConstants.COLLAPSE.equals(visibility)) {
            rPr.addNewVanish();
        }
    }

    /**
     * ????????????
     *
     * @param pictureData ???????????????
     * @param pictureType ????????????
     * @param filename ?????????
     * @param width ??????
     * @param height ??????
     * @param svgData SVG??????
     */
    public void renderPicture(InputStream pictureData, int pictureType, String filename, int width, int height, byte[] svgData)
            throws IOException, InvalidFormatException {
        CTR ctr = newRun();

        currentRun.addPicture(pictureData, pictureType, filename, width, height);
        CTR r = currentRun.getCTR();

        CTDrawing drawing = null;
        if (r != ctr) {
            int lastDrawingIndex = r.sizeOfDrawingArray() - 1;
            drawing = r.getDrawingArray(lastDrawingIndex);
            ctr.setDrawingArray(new CTDrawing[]{drawing});
            r.removeDrawing(lastDrawingIndex);
        } else if (svgData != null) {
            drawing = ctr.getDrawingArray(ctr.sizeOfDrawingArray() - 1);
        }

        if (svgData != null) {
            CTInline[] inlineArray = drawing.getInlineArray();
            if (inlineArray.length > 0) {
                CTInline ctInline = inlineArray[0];
                String svgRelId = getXWPFDocument().addPictureData(svgData, SVGPictureData.PICTURE_TYPE_SVG);
                CTNonVisualGraphicFrameProperties properties = ctInline.isSetCNvGraphicFramePr()
                        ? ctInline.getCNvGraphicFramePr() : ctInline.addNewCNvGraphicFramePr();
                CTGraphicalObjectFrameLocking frameLocking = properties.isSetGraphicFrameLocks()
                        ? properties.getGraphicFrameLocks() : properties.addNewGraphicFrameLocks();
                frameLocking.setNoChangeAspect(true);

                XmlCursor xmlCursor = ctInline.getGraphic().getGraphicData().newCursor();
                if (xmlCursor.toFirstChild()) {
                    CTPicture ctPicture = (CTPicture) xmlCursor.getObject();
                    CTBlip blip = ctPicture.getBlipFill().getBlip();
                    if (blip != null) {
                        CTOfficeArtExtensionList extList = blip.isSetExtLst() ? blip.getExtLst() : blip.addNewExtLst();
                        CTOfficeArtExtension svgBitmap = extList.addNewExt();
                        svgBitmap.setUri(SVGRelation.SVG_URI);
                        XmlCursor cur = svgBitmap.newCursor();
                        cur.toEndToken();
                        cur.beginElement(SVGRelation.SVG_QNAME);
                        cur.insertNamespace(SVGRelation.SVG_PREFIX, SVGRelation.MS_SVG_NS);
                        cur.insertAttributeWithValue(SVGRelation.EMBED_TAG, svgRelId);
                        cur.dispose();
                    }
                }
                xmlCursor.dispose();
            }
        }
    }

    /**
     * ??????????????????EMU
     *
     * @param length ??????
     * @return EMU
     */
    public int lengthToEMU(CSSLength length) {
        if (!length.isValid()) {
            throw new UnsupportedOperationException("Invalid CSS length");
        }
        if (!length.getUnit().isRelative()) {
            return length.toEMU();
        }
        double emu;
        switch (length.getUnit()) {
            case REM:
                emu = length.unitValue() * getDefaultFontSize().toEMU();
                break;
            case EM:
                emu = length.unitValue() * getInheritedFontSizeInHalfPoints() * Units.EMU_PER_POINT / 2;
                break;
            case VW:
                emu = length.unitValue() * getPageWidth().toEMU();
                break;
            case VH:
                emu = length.unitValue() * getPageHeight().toEMU();
                break;
            case VMIN:
                emu = length.unitValue() * Math.min(getPageWidth().toEMU(), getPageHeight().toEMU());
                break;
            case VMAX:
                emu = length.unitValue() * Math.max(getPageWidth().toEMU(), getPageHeight().toEMU());
                break;
            // Unable to determine the use of width or height as a relative length for percent unit
            default:
                throw new UnsupportedOperationException("Can not convert to EMU with length: " + length);
        }
        return (int) Math.rint(emu);
    }

    public NumberingContext getNumberingContext() {
        return numberingContext;
    }

    public CSSLength getPageWidth() {
        return pageWidth;
    }

    public CSSLength getPageHeight() {
        return pageHeight;
    }

    public CSSLength getMarginTop() {
        return marginTop;
    }

    public CSSLength getMarginRight() {
        return marginRight;
    }

    public CSSLength getMarginBottom() {
        return marginBottom;
    }

    public CSSLength getMarginLeft() {
        return marginLeft;
    }

    public int getAvailablePageWidth() {
        return availablePageWidth;
    }

    public int getAvailablePageHeight() {
        return availablePageHeight;
    }

    public XWPFRun getCurrentRun() {
        return currentRun;
    }

    public String getGlobalFont() {
        return globalFont;
    }

    public BigInteger getGlobalFontSize() {
        return globalFontSize;
    }

    public void setGlobalFont(String globalFont) {
        this.globalFont = globalFont;
    }

    public void setGlobalFontSize(BigInteger globalFontSize) {
        this.globalFontSize = globalFontSize;
    }

    public boolean isBlocked() {
        return blockLevel > 0;
    }

    public void incrementBlockLevel() {
        blockLevel++;
    }

    public void decrementBlockLevel() {
        blockLevel--;
    }
}

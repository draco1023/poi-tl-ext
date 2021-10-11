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
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
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
import org.jsoup.internal.StringUtil;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STThemeColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalAlignRun;

import javax.xml.namespace.QName;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.LinkedList;

/**
 * HTML字符串渲染上下文
 *
 * @author Draco
 * @since 2021-02-08
 */
public class HtmlRenderContext extends RenderContext<String> {
    private static final QName R_QNAME = new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "r");
    private static final QName HYPERLINK_QNAME = new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "hyperlink");

    /**
     * 默认字号 小四 12pt 16px
     */
    private static final CSSLength DEFAULT_FONT_SIZE = new CSSLength(12, CSSLengthUnit.PT);
    /**
     * 默认超链接颜色
     */
    private static final String DEFAULT_HYPERLINK_COLOR = "0563C1";

    /**
     * 最近的IBodyElement栈，主要用于渲染HTML表格时IBodyElement元素的切换
     */
    private LinkedList<IBodyElement> closestBodyStack = new LinkedList<>();

    /**
     * 父容器（子元素通常为段落/表格）栈，主要用于渲染HTML表格时父容器的切换
     */
    private LinkedList<IBody> ancestors = new LinkedList<>();

    /**
     * 行内样式栈，最近声明的样式最先生效
     */
    private LinkedList<InlineStyle> inlineStyles = new LinkedList<>();

    /**
     * 字号栈，一些相对大小的字号值将被进行换算
     */
    private LinkedList<Integer> fontSizesInHalfPoints = new LinkedList<>();

    /**
     * 列表上下文，用于处理嵌套列表
     */
    private final NumberingContext numberingContext;
    /**
     * 默认字号
     */
    private final CSSLength defaultFontSize;
    /**
     * 页面宽度
     */
    private final CSSLength pageWidth;
    /**
     * 页面高度
     */
    private final CSSLength pageHeight;
    /**
     * 页面顶部边距
     */
    private final CSSLength marginTop;
    /**
     * 页面右侧边距
     */
    private final CSSLength marginRight;
    /**
     * 页面底部边距
     */
    private final CSSLength marginBottom;
    /**
     * 页面左侧边距
     */
    private final CSSLength marginLeft;
    /**
     * 可用页面宽度
     */
    private final int availablePageWidth;
    /**
     * 可用页面高度
     */
    private final int availablePageHeight;

    /**
     * 当前Run元素，可能为超链接
     */
    private XWPFRun currentRun;

    /**
     * 全局字体，声明后所有样式中的字体将失效
     */
    private String globalFont;
    /**
     * 全局字号，声明后所有样式中的字号将失效
     */
    private BigInteger globalFontSize;

    /**
     * 块状元素深度计数器
     */
    private int blockLevel;

    /**
     * 构造方法
     *
     * @param context 原始渲染上下文
     */
    public HtmlRenderContext(RenderContext<String> context) {
        super(context.getEleTemplate(), context.getData(), context.getTemplate());
        numberingContext = new NumberingContext(getXWPFDocument());

        CTSectPr sectPr = getXWPFDocument().getDocument().getBody().getSectPr();
        CTPageSz pgSz = sectPr.getPgSz();
        // 页面尺寸单位是twip
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
     * 父容器入栈
     *
     * @param body 父容器
     */
    public void pushContainer(IBody body) {
        ancestors.push(body);
    }

    /**
     * 父容器出栈
     */
    public void popContainer() {
        ancestors.pop();
    }

    /**
     * 获取最近位置的内容元素（段落或表格）
     *
     * @return 最近位置的内容元素
     */
    public IBodyElement getClosestBody() {
        IBodyElement body = closestBodyStack.peek();
        return body == null ? (IBodyElement) getRun().getParent() : body;
    }

    /**
     * 内容元素（段落或表格）入栈
     *
     * @param body 内容元素（段落或表格）
     */
    public void pushClosestBody(IBodyElement body) {
        closestBodyStack.push(body);
    }

    /**
     * 内容元素（段落或表格）出栈
     */
    public void popClosestBody() {
        closestBodyStack.pop();
    }

    /**
     * 替换最近位置的内容元素（段落或表格）
     *
     * @param body 内容元素（段落或表格）
     */
    public void replaceClosestBody(IBodyElement body) {
        if (!closestBodyStack.isEmpty()) {
            closestBodyStack.pop();
        }
        closestBodyStack.push(body);
    }

    /**
     * 获取最近的段落，如果当前最近位置的内容元素是表格，则创建一个与之平级的段落
     *
     * @return 最近的段落
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
                XWPFParagraph paragraph = getContainer().insertNewParagraph(xmlCursor);
                replaceClosestBody(paragraph);
                xmlCursor.dispose();
                return paragraph;
        }
        throw new IllegalStateException("Impossible");
    }

    /**
     * 开始渲染超链接
     *
     * @param uri 链接地址
     */
    public void startHyperlink(String uri) {
        if (isBlocked()) {
            currentRun = getClosestParagraph().createHyperlinkRun(uri);
        } else {
            // 在占位符之前插入超链接
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
     * 结束渲染超链接
     */
    public void endHyperlink() {
        currentRun = null;
    }

    /**
     * 新建CTR
     *
     * @return CTR
     */
    public CTR newRun() {
        // 超链接虽然不是段落，但是内部可以容纳多个run
        if (currentRun instanceof XWPFHyperlinkRun) {
            XmlCursor xmlCursor = currentRun.getCTR().newCursor();
            CTR ctr;
            if (xmlCursor.toFirstChild()) {
                xmlCursor.dispose();
                ctr = ((XWPFHyperlinkRun) currentRun).getCTHyperlink().addNewR();
            } else {
                // run没有内容则直接复用
                ctr = currentRun.getCTR();
            }
            // 默认链接样式
            initHyperlinkStyle(ctr);

            return ctr;
        }
        // 考虑到样式可能不一致，总是创建新的run
        if (isBlocked()) {
            currentRun = getClosestParagraph().createRun();
        } else {
            // 在占位符之前插入run
            XmlCursor xmlCursor = getRun().getCTR().newCursor();
            xmlCursor.insertElement(R_QNAME);
            xmlCursor.toPrevSibling();
            CTR ctr = (CTR) xmlCursor.getObject();
            currentRun = new XWPFRun(ctr, getRun().getParent());
        }
        return currentRun.getCTR();
    }

    /**
     * 初始化超链接样式
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
     * 获取最近的表格，仅可在渲染表格及其内部元素的时候使用
     *
     * @return 最近的表格
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
     * 行内样式入栈
     *
     * @param inlineStyle 样式声明
     * @param block 是否为块状元素
     */
    public void pushInlineStyle(CSSStyleDeclarationImpl inlineStyle, boolean block) {
        String newFontSize = inlineStyle.getFontSize();
        // 默认值表示未声明字号
        int fontSize = Integer.MIN_VALUE;
        if (StringUtils.isNotBlank(newFontSize)) {
            NamedFontSize namedFontSize = NamedFontSize.of(newFontSize);
            if (namedFontSize != null) {
                // 固定名称的字号
                fontSize = namedFontSize.getSize().toHalfPoints();
            } else if (HtmlConstants.SMALLER.equalsIgnoreCase(newFontSize)) {
                // 相对小一号
                int inheritedFontSize = getInheritedFontSizeInHalfPoints();
                fontSize = RenderUtils.smallerFontSizeInHalfPoints(inheritedFontSize);
            } else if (HtmlConstants.LARGER.equalsIgnoreCase(newFontSize)) {
                // 相对大一号
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
     * 行内样式出栈
     */
    public void popInlineStyle() {
        fontSizesInHalfPoints.pop();
        inlineStyles.pop();
    }

    /**
     * 当前元素的样式声明，在HTML元素渲染开始时立即调用才可得到正确的声明，因为在解析的过程中可能会动态插入样式
     *
     * @return 当前元素的样式声明
     */
    public CSSStyleDeclarationImpl currentElementStyle() {
        InlineStyle inlineStyle = inlineStyles.peek();
        return inlineStyle == null ? CSSStyleUtils.EMPTY_STYLE : inlineStyle.getDeclaration();
    }

    /**
     * 获取样式值，将被转换为小写
     *
     * @param property 样式名称
     * @return 样式值，未声明时返回空字符串
     */
    public String getPropertyValue(String property) {
        return getPropertyValue(property, false);
    }

    /**
     * 获取样式值，将被转换为小写
     *
     * @param property 样式名称
     * @param inlineOnly 是否仅获取行内元素样式
     * @return 样式值，未声明时返回空字符串
     */
    public String getPropertyValue(String property, boolean inlineOnly) {
        return getPropertyValue(property, false, inlineOnly);
    }

    /**
     * 获取样式值
     *
     * @param property 样式名称
     * @param caseSensitive 是否大小写无关，如果无关则将转换为小写，否则保留原始值
     * @param inlineOnly 是否仅获取行内元素样式
     * @return 样式值，未声明时返回空字符串
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
     * @return Word中设置的默认字号
     */
    public CSSLength getDefaultFontSize() {
        return defaultFontSize;
    }

    /**
     * @return 获取当前元素继承的字号，以“半点”为单位
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
     * @return 父容器的可用宽度，以EMU为单位
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
     * 考虑约束计算长度，以EMU为单位
     *
     * @param length 长度声明
     * @param maxLength 最大长度声明
     * @param naturalEMU 原始长度
     * @param parentEMU 父容器长度
     * @return 以EMU为单位的长度值
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
     * 考虑约束计算长度，以EMU为单位
     *
     * @param cssLength 长度声明
     * @param naturalEMU 原始长度
     * @param parentEMU 父容器长度
     * @return 以EMU为单位的长度值
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
     * 渲染文本
     *
     * @param text 文本
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

        // 应用样式
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
     * 应用文本样式
     *
     * @param ctr CTR
     */
    private void applyTextStyle(CTR ctr) {
        CTRPr rPr = RenderUtils.getRPr(ctr);

        // 字体，如果声明了全局字体则忽略样式声明
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

        // 字号
        if (globalFontSize == null) {
            String fontSize = getPropertyValue(HtmlConstants.CSS_FONT_SIZE);
            if (StringUtils.isNotBlank(fontSize)) {
                int sz = getInheritedFontSizeInHalfPoints();
                rPr.addNewSz().setVal(BigInteger.valueOf(sz));
            }
        } else {
            // 如果定义了全局字号则忽略样式声明
            rPr.addNewSz().setVal(globalFontSize);
        }

        // 加粗
        String fontWeight = getPropertyValue(HtmlConstants.CSS_FONT_WEIGHT);
        if (fontWeight.contains(HtmlConstants.BOLD)) {
            rPr.addNewB();
        } else if (NumberUtils.isParsable(fontWeight) && Float.parseFloat(fontWeight) > 500) {
            rPr.addNewB();
        }

        // 斜体
        String fontStyle = getPropertyValue(HtmlConstants.CSS_FONT_STYLE);
        if (HtmlConstants.ITALIC.equals(fontStyle) || HtmlConstants.OBLIQUE.equals(fontStyle)) {
            rPr.addNewI();
        }

        // FIXME cssparser目前不支持以空格为分隔符的rgb颜色，仅支持逗号分隔符
        // 颜色
        String color = getPropertyValue(HtmlConstants.CSS_COLOR);
        if (StringUtils.isNotBlank(color)) {
            String hex = Colors.fromStyle(color);
            RenderUtils.getColor(rPr).setVal(hex);
        }

        String caps = getPropertyValue(HtmlConstants.CSS_FONT_VARIANT_CAPS);
        if (HtmlConstants.SMALL_CAPS.equals(caps)) {
            rPr.addNewSmallCaps();
        }

        // 中划线/下划线
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

        // 上下标
        String verticalAlign = getPropertyValue(HtmlConstants.CSS_VERTICAL_ALIGN);
        if (HtmlConstants.SUPER.equals(verticalAlign)) {
            rPr.addNewVertAlign().setVal(STVerticalAlignRun.SUPERSCRIPT);
        } else if (HtmlConstants.SUB.equals(verticalAlign)) {
            rPr.addNewVertAlign().setVal(STVerticalAlignRun.SUBSCRIPT);
        }

        // FIXME 段落边框与行内边框分离，行内只有全边框，段落分四边

        // 可见性
        String visibility = getPropertyValue(HtmlConstants.CSS_VISIBILITY);
        if (HtmlConstants.HIDDEN.equals(visibility) || HtmlConstants.COLLAPSE.equals(visibility)) {
            rPr.addNewVanish();
        }
    }

    /**
     * 渲染图片
     *
     * @param pictureData 图片数据流
     * @param pictureType 图片类型
     * @param filename 文件名
     * @param width 宽度
     * @param height 高度
     */
    public void renderPicture(InputStream pictureData, int pictureType, String filename, int width, int height)
            throws IOException, InvalidFormatException {
        CTR ctr = newRun();

        currentRun.addPicture(pictureData, pictureType, filename, width, height);
        CTR r = currentRun.getCTR();
        if (r != ctr) {
            int lastDrawingIndex = r.sizeOfDrawingArray() - 1;
            CTDrawing drawing = r.getDrawingArray(lastDrawingIndex);
            ctr.setDrawingArray(new CTDrawing[]{drawing});
            r.removeDrawing(lastDrawingIndex);
        }
    }

    /**
     * 将长度换算为EMU
     *
     * @param length 长度
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

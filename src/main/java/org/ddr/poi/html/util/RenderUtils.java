package org.ddr.poi.html.util;

import com.steadystate.css.dom.CSSStyleDeclarationImpl;
import com.steadystate.css.dom.CSSValueImpl;
import com.steadystate.css.dom.Property;
import com.steadystate.css.parser.CSSOMParser;
import com.steadystate.css.parser.SACParserCSS3;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BodyType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlString;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPBdr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;
import org.w3c.css.sac.InputSource;
import org.w3c.dom.css.CSSValue;

import javax.xml.namespace.QName;
import java.io.IOException;
import java.io.StringReader;
import java.math.BigInteger;
import java.util.function.Function;

/**
 * 渲染相关的工具类
 *
 * @author Draco
 * @since 2021-02-08
 */
public class RenderUtils {
    /**
     * CSS解析器
     */
    public static final CSSOMParser CSS_PARSER = new CSSOMParser(new SACParserCSS3());
    /**
     * 空样式
     */
    public static final CSSStyleDeclarationImpl EMPTY_STYLE = new CSSStyleDeclarationImpl();
    /**
     * Word中字号下拉列表对应的值
     */
    public static final int[] FONT_SIZE_IN_HALF_POINTS = {10, 11, 13, 15, 18, 21, 24, 28, 30, 32, 36, 44, 48, 52, 72, 84, 96, 144};
    /**
     * 边框宽度每像素对应值
     */
    public static final int BORDER_WIDTH_PER_PX = 4;
    /**
     * 边框纵向间距
     */
    public static final BigInteger VERTICAL_SPACE = BigInteger.ONE;
    /**
     * 边框横向间距
     */
    public static final BigInteger HORIZONTAL_SPACE = BigInteger.valueOf(4);
    /**
     * 最小边框宽度
     */
    public static final long MIN_BORDER_WIDTH = 2;
    /**
     * 最大边框宽度
     */
    public static final long MAX_BORDER_WIDTH = 96;

    /**
     * 表格单元格边距
     */
    public static final int TABLE_CELL_MARGIN = 108;

    /**
     * 解析行内样式
     *
     * @param inlineStyle 行内样式声明
     * @return 样式
     */
    public static CSSStyleDeclarationImpl parse(String inlineStyle) {
        try (StringReader sr = new StringReader(inlineStyle)) {
            return (CSSStyleDeclarationImpl) CSS_PARSER.parseStyleDeclaration(new InputSource(sr));
        } catch (IOException ignored) {
            return RenderUtils.EMPTY_STYLE;
        }
    }

    /**
     * 解析样式值
     *
     * @param value 样式值字符串
     * @return 样式值
     */
    public static CSSValue parseValue(String value) {
        try (StringReader sr = new StringReader(value)) {
            return CSS_PARSER.parsePropertyValue(new InputSource(sr));
        } catch (IOException ignored) {
            return new CSSValueImpl();
        }
    }

    /**
     * 将样式键值对转换为样式属性
     *
     * @param key 样式名称
     * @param value 样式值
     * @return 样式属性
     */
    public static Property newProperty(String key, String value) {
        return new Property(key, parseValue(value), false);
    }

    /**
     * 文本对齐值映射
     *
     * @param textAlign 文本对齐样式值
     * @return Word文本对齐枚举
     */
    public static ParagraphAlignment align(String textAlign) {
        if (StringUtils.isBlank(textAlign)) {
            return null;
        }
        switch (textAlign.toLowerCase()) {
            case HtmlConstants.START:
            case HtmlConstants.LEFT:
                return ParagraphAlignment.LEFT;
            case HtmlConstants.END:
            case HtmlConstants.RIGHT:
                return ParagraphAlignment.RIGHT;
            case HtmlConstants.CENTER:
                return ParagraphAlignment.CENTER;
            case HtmlConstants.JUSTIFY:
            case HtmlConstants.JUSTIFY_ALL:
                return ParagraphAlignment.BOTH;
            default:
                return null;
        }
    }

    /**
     * 下划线样式映射
     *
     * @param textDecorationStyle 下划线样式值
     * @return Word下划线样式
     */
    public static STUnderline.Enum underline(String textDecorationStyle) {
        switch (textDecorationStyle) {
//            case HtmlConstants.SOLID:
//                return STUnderline.SINGLE;
            case HtmlConstants.DOUBLE:
                return STUnderline.DOUBLE;
            case HtmlConstants.DOTTED:
                return STUnderline.DOTTED;
            case HtmlConstants.DASHED:
                return STUnderline.DASH;
            case HtmlConstants.WAVY:
                return STUnderline.WAVE;
            default:
                return STUnderline.SINGLE;
        }
    }

    public static CTPPr getPPr(CTStyle ctStyle) {
        return ctStyle.isSetPPr() ? ctStyle.getPPr() : ctStyle.addNewPPr();
    }

    public static CTPPr getPPr(CTP ctp) {
        return ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
    }

    public static CTPBdr getPBdr(CTPPr pr) {
        return pr.isSetPBdr() ? pr.getPBdr() : pr.addNewPBdr();
    }

    public static CTJc getJc(CTPPr pr) {
        return pr.isSetJc() ? pr.getJc() : pr.addNewJc();
    }

    public static CTRPr getRPr(CTR ctr) {
        return ctr.isSetRPr() ? ctr.getRPr() : ctr.addNewRPr();
    }

    public static CTTcPr getTcPr(CTTc tc) {
        return tc.isSetTcPr() ? tc.getTcPr() : tc.addNewTcPr();
    }

    public static CTShd getShd(CTPPr pPr) {
        return pPr.isSetShd() ? pPr.getShd() : pPr.addNewShd();
    }

    public static CTInd getInd(CTPPr pPr) {
        return pPr.isSetInd() ? pPr.getInd() : pPr.addNewInd();
    }

    public static CTColor getColor(CTRPr rPr) {
        return rPr.isSetColor() ? rPr.getColor() : rPr.addNewColor();
    }

    public static CTUnderline getUnderline(CTRPr rPr) {
        return rPr.isSetU() ? rPr.getU() : rPr.addNewU();
    }

    /**
     * 获取父容器的可用宽度，以EMU为单位
     *
     * @param body 父容器
     * @return 可用宽度
     */
    public static int getAvailableWidthInEMU(IBody body) {
        if (body.getPartType() == BodyType.DOCUMENT) {
            XWPFDocument document = (XWPFDocument) body;
            CTSectPr sectPr = document.getDocument().getBody().getSectPr();
            int availableWidth = sectPr.getPgSz().getW().intValue()
                    - sectPr.getPgMar().getLeft().intValue() - sectPr.getPgMar().getRight().intValue();
            return Units.TwipsToEMU((short) availableWidth);

        } else if (body.getPartType() == BodyType.TABLECELL) {
            XWPFTableCell tableCell = ((XWPFTableCell) body);
            CTTblWidth tcW = tableCell.getCTTc().getTcPr().getTcW();
            if (TableWidthType.DXA.getStWidthType().equals(tcW.getType())) {
                int availableWidth = tcW.getW().intValue() - TABLE_CELL_MARGIN * 2;
                return availableWidth > 0 ? Units.TwipsToEMU((short) availableWidth) : 0;
            } else if (TableWidthType.PCT.getStWidthType().equals(tcW.getType())) {
                CTTblWidth tblW = tableCell.getTableRow().getTable().getCTTbl().getTblPr().getTblW();
                if (TableWidthType.DXA.getStWidthType().equals(tblW.getType())) {
                    int availableWidth = tblW.getW().intValue() * tcW.getW().intValue() / 5000 - TABLE_CELL_MARGIN * 2;
                    return availableWidth > 0 ? Units.TwipsToEMU((short) availableWidth) : 0;
                } else if (TableWidthType.NIL.getStWidthType().equals(tblW.getType())) {
                    return 0;
                } else {
                    return Integer.MAX_VALUE;
                }
            } else if (TableWidthType.NIL.getStWidthType().equals(tcW.getType())) {
                return 0;
            } else {
                return Integer.MAX_VALUE;
            }

        } else {
            throw new UnsupportedOperationException("Get bounds of " + body.getPartType() + " is not supported yet");
        }
    }

    /**
     * 应用段落样式
     *
     * @param context 渲染上下文
     * @param paragraph 段落
     * @param cssStyleDeclaration CSS样式声明
     */
    public static void paragraphStyle(HtmlRenderContext context, XWPFParagraph paragraph, CSSStyleDeclarationImpl cssStyleDeclaration) {
        if (EMPTY_STYLE.equals(cssStyleDeclaration)) {
            return;
        }

        // alignment
        ParagraphAlignment align = align(cssStyleDeclaration.getTextAlign());
        if (align != null) {
            paragraph.setAlignment(align);
        }

        // border
        setBorder(paragraph, cssStyleDeclaration, HtmlConstants.CSS_BORDER_TOP_STYLE, HtmlConstants.CSS_BORDER_TOP_WIDTH,
                HtmlConstants.CSS_BORDER_TOP_COLOR, RenderUtils::getTop, VERTICAL_SPACE);
        setBorder(paragraph, cssStyleDeclaration, HtmlConstants.CSS_BORDER_RIGHT_STYLE, HtmlConstants.CSS_BORDER_RIGHT_WIDTH,
                HtmlConstants.CSS_BORDER_RIGHT_COLOR, RenderUtils::getRight, HORIZONTAL_SPACE);
        setBorder(paragraph, cssStyleDeclaration, HtmlConstants.CSS_BORDER_BOTTOM_STYLE, HtmlConstants.CSS_BORDER_BOTTOM_WIDTH,
                HtmlConstants.CSS_BORDER_BOTTOM_COLOR, RenderUtils::getBottom, VERTICAL_SPACE);
        setBorder(paragraph, cssStyleDeclaration, HtmlConstants.CSS_BORDER_LEFT_STYLE, HtmlConstants.CSS_BORDER_LEFT_WIDTH,
                HtmlConstants.CSS_BORDER_LEFT_COLOR, RenderUtils::getLeft, HORIZONTAL_SPACE);

        // TODO spacing

        // indent
        CSSValueImpl textIndent = (CSSValueImpl) cssStyleDeclaration.getPropertyCSSValue(HtmlConstants.CSS_TEXT_INDENT);
        if (textIndent != null) {
            int length = textIndent.getLength();
            if (length == 0) {
                indent(context, paragraph, textIndent.getCssText());
            } else {
                for (int i = 0; i < length; i++) {
                    CSSValue item = textIndent.item(i);
                    boolean indented = indent(context, paragraph, item.getCssText());
                    if (indented) {
                        break;
                    }
                }
            }
        }

        // background
        String backgroundColor = cssStyleDeclaration.getBackgroundColor();
        if (StringUtils.isNotBlank(backgroundColor)) {
            String color = Colors.fromStyle(backgroundColor, null);
            if (color != null) {
                CTPPr pPr = getPPr(paragraph.getCTP());
                CTShd shd = getShd(pPr);
                shd.setFill(color);
            }
        }
    }

    /**
     * 段落首行缩进
     *
     * @param context 渲染上下文
     * @param paragraph 段落
     * @param style 缩进样式值
     * @return 是否进行了缩进
     */
    private static boolean indent(HtmlRenderContext context, XWPFParagraph paragraph, String style) {
        CSSLength cssLength = CSSLength.of(style.toLowerCase());
        if (cssLength.isValid() && cssLength.getValue() > 0) {
            CTPPr pPr = getPPr(paragraph.getCTP());
            CTInd ind = getInd(pPr);
            double indent;
            if (cssLength.isPercent()) {
                indent = context.getAvailableWidthInEMU() * cssLength.unitValue() / CSSLengthUnit.TWIP.absoluteFactor();
            } else {
                indent = context.lengthToEMU(cssLength) / CSSLengthUnit.TWIP.absoluteFactor();
            }
            ind.setFirstLine(BigInteger.valueOf(Math.round(indent)));
            return true;
        }
        return false;
    }

    /**
     * 设置段落边框样式
     *
     * @param paragraph 段落
     * @param cssStyleDeclaration CSS边框样式声明
     * @param styleProperty CSS边框属性名称
     * @param widthProperty CSS边框宽度名称
     * @param colorProperty CSS边框颜色名称
     * @param getter 获取边框对象的方式
     * @param space 边框间距
     */
    private static void setBorder(XWPFParagraph paragraph, CSSStyleDeclarationImpl cssStyleDeclaration,
                                  String styleProperty, String widthProperty, String colorProperty,
                                  Function<CTPBdr, CTBorder> getter, BigInteger space) {
        String borderStyle = cssStyleDeclaration.getPropertyValue(styleProperty);
        STBorder.Enum style = borderStyle(borderStyle);
        String borderWidth = cssStyleDeclaration.getPropertyValue(widthProperty);
        CSSLength width = CSSLength.of(borderWidth);
        if (style != null && (!width.isValid() || width.getValue() > 0)) {
            CTPPr pPr = getPPr(paragraph.getCTP());
            CTPBdr pBdr = getPBdr(pPr);
            CTBorder border = getter.apply(pBdr);
            border.setVal(style);

            String borderColor = cssStyleDeclaration.getPropertyValue(colorProperty);
            String color = Colors.fromStyle(borderColor);
            border.setColor(color);

            if (width.isValid() && !width.isPercent()) {
                long widthValue = Math.round(width.to(CSSLengthUnit.PX).getValue()) * BORDER_WIDTH_PER_PX;
                if (widthValue < MIN_BORDER_WIDTH) {
                    widthValue = MIN_BORDER_WIDTH;
                } else if (widthValue > MAX_BORDER_WIDTH) {
                    widthValue = MAX_BORDER_WIDTH;
                }
                border.setSz(BigInteger.valueOf(widthValue));
            } else {
                border.setSz(BigInteger.valueOf(BORDER_WIDTH_PER_PX));
            }

            border.setSpace(space);
        }
    }

    /**
     * 边框样式映射
     *
     * @param style CSS边框样式值
     * @return Word边框样式
     */
    private static STBorder.Enum borderStyle(String style) {
        if (StringUtils.isBlank(style)) {
            return null;
        }

        switch (style.toLowerCase()) {
            case HtmlConstants.DOTTED:
                return STBorder.DOTTED;
            case HtmlConstants.DASHED:
                return STBorder.DASHED;
            case HtmlConstants.SOLID:
                return STBorder.SINGLE;
            case HtmlConstants.DOUBLE:
                return STBorder.DOUBLE;
            case HtmlConstants.GROOVE:
            case HtmlConstants.INSET:
                return STBorder.INSET;
            case HtmlConstants.RIDGE:
            case HtmlConstants.OUTSET:
                return STBorder.OUTSET;
            default:
                return null;
        }
    }

    public static CTBorder getTop(CTPBdr pBdr) {
        return pBdr.isSetTop() ? pBdr.getTop() : pBdr.addNewTop();
    }

    public static CTBorder getRight(CTPBdr pBdr) {
        return pBdr.isSetRight() ? pBdr.getRight() : pBdr.addNewRight();
    }

    public static CTBorder getBottom(CTPBdr pBdr) {
        return pBdr.isSetBottom() ? pBdr.getBottom() : pBdr.addNewBottom();
    }

    public static CTBorder getLeft(CTPBdr pBdr) {
        return pBdr.isSetLeft() ? pBdr.getLeft() : pBdr.addNewLeft();
    }

    /**
     * 获取小一号字号
     *
     * @param inheritedSizeInHalfPoints 当前字号
     * @return 字号
     */
    public static int smallerFontSizeInHalfPoints(int inheritedSizeInHalfPoints) {
        for (int i = FONT_SIZE_IN_HALF_POINTS.length - 1; i >= 0; i--) {
            int s = FONT_SIZE_IN_HALF_POINTS[i];
            if (s < inheritedSizeInHalfPoints) {
                return s;
            }
        }
        return FONT_SIZE_IN_HALF_POINTS[0];
    }

    /**
     * 获取大一号字号
     *
     * @param inheritedSizeInHalfPoints 当前字号
     * @return 字号
     */
    public static int largerFontSizeInHalfPoints(int inheritedSizeInHalfPoints) {
        for (int s : FONT_SIZE_IN_HALF_POINTS) {
            if (s > inheritedSizeInHalfPoints) {
                return s;
            }
        }
        return FONT_SIZE_IN_HALF_POINTS[FONT_SIZE_IN_HALF_POINTS.length - 1];
    }

    /**
     * EMU转twip
     *
     * @see Units#TwipsToEMU
     */
    public static int emuToTwips(int emu) {
        return (int) (emu * 20L / Units.EMU_PER_POINT);
    }

    /**
     * @see XWPFRun#preserveSpaces(XmlString)
     */
    public static void preserveSpaces(XmlString xs) {
        String text = xs.getStringValue();
        if (text != null && text.length() >= 1
                && (Character.isWhitespace(text.charAt(0)) || Character.isWhitespace(text.charAt(text.length() - 1)))) {
            XmlCursor c = xs.newCursor();
            c.toNextToken();
            c.insertAttributeWithValue(new QName("http://www.w3.org/XML/1998/namespace", "space"), "preserve");
            c.dispose();
        }
    }
}

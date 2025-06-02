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

package org.ddr.poi.html.util;

import com.steadystate.css.dom.CSSStyleDeclarationImpl;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ooxml.util.POIXMLUnits;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BodyType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TableRowAlign;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPoint2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPBdr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPrGeneral;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;

import java.math.BigInteger;
import java.util.List;
import java.util.function.Function;

/**
 * 渲染相关的工具类
 *
 * @author Draco
 * @since 2021-02-08
 */
public class RenderUtils {
    /**
     * Word中字号下拉列表对应的值
     */
    public static final int[] FONT_SIZE_IN_HALF_POINTS = {10, 11, 13, 15, 18, 21, 24, 28, 30, 32, 36, 44, 48, 52, 72, 84, 96, 144};
    /**
     * 边框宽度每像素对应值
     */
    public static final int BORDER_WIDTH_PER_PX = 4;
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
     * 段落行距系数
     */
    public static final int SPACING_FACTOR = 240;
    /**
     * 默认页面宽度 A4 portrait
     */
    public static final long A4_WIDTH = 11906;
    /**
     * 默认页面高度 A4 portrait
     */
    public static final long A4_HEIGHT = 16838;

    // Defined in [ Word 2010 look.dotx ]
    /**
     * 默认顶边距
     */
    public static final long DEFAULT_TOP_MARGIN = 1440;
    /**
     * 默认底边距
     */
    public static final long DEFAULT_BOTTOM_MARGIN = 1440;
    /**
     * 默认左边距
     */
    public static final long DEFAULT_LEFT_MARGIN = 1440;
    /**
     * 默认右边距
     */
    public static final long DEFAULT_RIGHT_MARGIN = 1440;

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

    public static CTPPrGeneral getPPr(CTStyle ctStyle) {
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

    public static CTTcMar getTcMar(CTTcPr tcPr) {
        return tcPr.isSetTcMar() ? tcPr.getTcMar() : tcPr.addNewTcMar();
    }

    public static CTTcMar getTcMar(XWPFTableCell cell) {
        CTTcPr tcPr = getTcPr(cell.getCTTc());
        return getTcMar(tcPr);
    }

    public static CTShd getShd(CTPPr pPr) {
        return pPr.isSetShd() ? pPr.getShd() : pPr.addNewShd();
    }

    public static CTInd getInd(CTPPr pPr) {
        return pPr.isSetInd() ? pPr.getInd() : pPr.addNewInd();
    }

    public static CTInd getInd(XWPFParagraph paragraph) {
        CTPPr pPr = getPPr(paragraph.getCTP());
        return getInd(pPr);
    }

    public static CTSpacing getSpacing(CTPPr pPr) {
        return pPr.isSetSpacing() ? pPr.getSpacing() : pPr.addNewSpacing();
    }

    public static CTSpacing getSpacing(XWPFParagraph paragraph) {
        CTPPr pPr = getPPr(paragraph.getCTP());
        return getSpacing(pPr);
    }

    public static CTColor getColor(CTRPr rPr) {
        List<CTColor> colorList = rPr.getColorList();
        return colorList.isEmpty() ? rPr.addNewColor() : colorList.get(0);
    }

    public static CTUnderline getUnderline(CTRPr rPr) {
        List<CTUnderline> uList = rPr.getUList();
        return uList.isEmpty() ? rPr.addNewU() : uList.get(0);
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
            long availableWidth = POIXMLUnits.parseLength(sectPr.getPgSz().xgetW())
                    - POIXMLUnits.parseLength(sectPr.getPgMar().xgetLeft())
                    - POIXMLUnits.parseLength(sectPr.getPgMar().xgetRight());
            return (int) availableWidth;

        } else if (body.getPartType() == BodyType.TABLECELL) {
            XWPFTableCell tableCell = ((XWPFTableCell) body);
            CTTblWidth tcW = tableCell.getCTTc().getTcPr().getTcW();
            if (TableWidthType.DXA.getStWidthType().equals(tcW.getType())) {
                long availableWidth = POIXMLUnits.parseLength(tcW.xgetW()) - twipsToEMU(TABLE_CELL_MARGIN) * 2;
                return availableWidth > 0 ? (int) availableWidth : 0;
            } else if (TableWidthType.PCT.getStWidthType().equals(tcW.getType())) {
                CTTblWidth tblW = tableCell.getTableRow().getTable().getCTTbl().getTblPr().getTblW();
                if (TableWidthType.DXA.getStWidthType().equals(tblW.getType())) {
                    String cellWidth = tcW.xgetW().getStringValue();
                    double percent;
                    if (cellWidth.endsWith("%")) { // 2011
                        percent = Double.parseDouble(cellWidth.replace("%", ""));
                    } else { // 2006
                        percent = Double.parseDouble(cellWidth) / 5000;
                    }
                    long availableWidth = Math.round(POIXMLUnits.parseLength(tblW.xgetW()) * percent)
                            - twipsToEMU(TABLE_CELL_MARGIN) * 2;
                    return availableWidth > 0 ? (int) availableWidth : 0;
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
        /* inheritable styles */
        // alignment
        ParagraphAlignment align = align(context.getPropertyValue(HtmlConstants.CSS_TEXT_ALIGN));
        if (align != null) {
            paragraph.setAlignment(align);
        }

        if (CSSStyleUtils.isEmpty(cssStyleDeclaration)) {
            return;
        }

        // border
        setBorder(context, paragraph, cssStyleDeclaration);

        // spacing
        setSpacing(context, paragraph, cssStyleDeclaration);

        // indent
        setIndentation(context, paragraph, cssStyleDeclaration);

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
     * 设置段落行距
     *
     * @param context 渲染上下文
     * @param paragraph 段落
     * @param cssStyleDeclaration CSS样式声明
     */
    private static void setSpacing(HtmlRenderContext context, XWPFParagraph paragraph,
                                   CSSStyleDeclarationImpl cssStyleDeclaration) {
        // margin-top
        CSSLength marginTop = CSSLength.of(cssStyleDeclaration.getMarginTop().toLowerCase());
        if (marginTop.isValid() && !marginTop.isPercent()) {
            getSpacing(paragraph).setBefore(BigInteger.valueOf(emuToTwips(context.lengthToEMU(marginTop))));
        }

        // margin-bottom
        CSSLength marginBottom = CSSLength.of(cssStyleDeclaration.getMarginBottom().toLowerCase());
        if (marginBottom.isValid() && !marginBottom.isPercent()) {
            getSpacing(paragraph).setAfter(BigInteger.valueOf(emuToTwips(context.lengthToEMU(marginBottom))));
        }

        // line-height
        String lineHeight = context.getPropertyValue(HtmlConstants.CSS_LINE_HEIGHT);
        if (StringUtils.isNotBlank(lineHeight)) {
            CSSLength cssLength = CSSLength.of(lineHeight);
            if (cssLength.isValid()) {
                if (cssLength.isPercent()) {
                    CTSpacing spacing = getSpacing(paragraph);
                    spacing.setLineRule(STLineSpacingRule.AUTO);
                    spacing.setLine(BigInteger.valueOf(Math.round(cssLength.unitValue() * SPACING_FACTOR)));
                } else if (cssLength.getValue() > 0) {
                    CTSpacing spacing = getSpacing(paragraph);
                    spacing.setLineRule(STLineSpacingRule.EXACT);
                    spacing.setLine(BigInteger.valueOf(emuToTwips(context.lengthToEMU(cssLength))));
                }
            } else if (NumberUtils.isParsable(lineHeight)) {
                double value = Double.parseDouble(lineHeight);
                if (value > 0) {
                    CTSpacing spacing = getSpacing(paragraph);
                    spacing.setLineRule(STLineSpacingRule.AUTO);
                    spacing.setLine(BigInteger.valueOf(Math.round(value * SPACING_FACTOR)));
                }
            }
        }
    }

    /**
     * 设置段落缩进
     *
     * @param context 渲染上下文
     * @param paragraph 段落
     * @param cssStyleDeclaration CSS样式声明
     */
    private static void setIndentation(HtmlRenderContext context, XWPFParagraph paragraph,
                                       CSSStyleDeclarationImpl cssStyleDeclaration) {
        // margin-left
        CSSLength marginLeft = CSSLength.of(cssStyleDeclaration.getMarginLeft().toLowerCase());
        if (marginLeft.isValid() && !marginLeft.isPercent()) {
            getInd(paragraph).setLeft(BigInteger.valueOf(emuToTwips(context.lengthToEMU(marginLeft))));
        }

        // margin-right
        CSSLength marginRight = CSSLength.of(cssStyleDeclaration.getMarginRight().toLowerCase());
        if (marginRight.isValid() && !marginRight.isPercent()) {
            getInd(paragraph).setRight(BigInteger.valueOf(emuToTwips(context.lengthToEMU(marginRight))));
        }

        // text-indent
        String textIndent = context.getPropertyValue(HtmlConstants.CSS_TEXT_INDENT).trim();
        if (textIndent.isEmpty()) {
            indent(context, paragraph, textIndent);
        } else {
            int nonSpace = -1;
            boolean indented = false;
            for (int i = 0; i < textIndent.length(); i++) {
                char c = textIndent.charAt(i);
                if (Character.isWhitespace(c)) {
                    if (nonSpace >= 0) {
                        indented = indent(context, paragraph, textIndent.substring(nonSpace, i));
                    }
                    nonSpace = -1;
                } else if (nonSpace < 0) {
                    nonSpace = i;
                }
                if (indented) {
                    break;
                }
            }
            if (nonSpace >= 0) {
                indent(context, paragraph, textIndent.substring(nonSpace));
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
     * @param context 上下文
     * @param xwpfElement 段落
     * @param cssStyleDeclaration CSS边框样式声明
     * @param styleProperty CSS边框属性名称
     * @param widthProperty CSS边框宽度名称
     * @param colorProperty CSS边框颜色名称
     * @param getter 获取边框对象的方式
     * @return 边框是否为none
     */
    private static boolean setBorder(HtmlRenderContext context, Object xwpfElement, CSSStyleDeclarationImpl cssStyleDeclaration,
                                     String styleProperty, String widthProperty, String colorProperty,
                                     Function<Object, CTBorder> getter) {
        String borderStyle = cssStyleDeclaration.getPropertyValue(styleProperty);
        STBorder.Enum style = borderStyle(borderStyle);
        String borderWidth = cssStyleDeclaration.getPropertyValue(widthProperty);
        CSSLength width = CSSLength.of(borderWidth);
        if (style != null && (!width.isValid() || width.getValue() > 0)) {
            CTBorder border = getter.apply(xwpfElement);
            border.setVal(style);

            String borderColor = cssStyleDeclaration.getPropertyValue(colorProperty);
            String color = Colors.TRANSPARENT.equals(borderColor) ? Colors.WHITE : Colors.fromStyle(borderColor);
            border.setColor(color);

            if (width.isValid() && !width.isPercent()) {
                long widthValue = (long) context.lengthToEMU(width) * BORDER_WIDTH_PER_PX / Units.EMU_PER_PIXEL;
                if (widthValue < MIN_BORDER_WIDTH) {
                    widthValue = MIN_BORDER_WIDTH;
                } else if (widthValue > MAX_BORDER_WIDTH) {
                    widthValue = MAX_BORDER_WIDTH;
                }
                border.setSz(BigInteger.valueOf(widthValue));
            } else {
                border.setSz(BigInteger.valueOf(BORDER_WIDTH_PER_PX));
            }
        }
        return style == STBorder.NONE || style == STBorder.NIL;
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
            case HtmlConstants.NONE:
                return STBorder.NONE;
            default:
                return null;
        }
    }

    public static CTTblBorders getTblBorders(CTTbl tbl) {
        CTTblPr tblPr = getTblPr(tbl);
        return getTblBorders(tblPr);
    }

    private static CTBorder getTop(Object e) {
        if (e instanceof XWPFParagraph) {
            XWPFParagraph paragraph = (XWPFParagraph) e;
            return getParagraphTop(paragraph.getCTP());
        } else if (e instanceof XWPFTable) {
            XWPFTable table = (XWPFTable) e;
            return getTableTop(table.getCTTbl());
        } else if (e instanceof XWPFTableCell) {
            XWPFTableCell cell = (XWPFTableCell) e;
            return getTableCellTop(cell.getCTTc());
        } else {
            throw new UnsupportedOperationException("Can not get top border of " + e.getClass().getName());
        }
    }

    public static CTBorder getParagraphTop(CTP paragraph) {
        CTPPr pPr = getPPr(paragraph);
        CTPBdr pBdr = getPBdr(pPr);
        return pBdr.isSetTop() ? pBdr.getTop() : pBdr.addNewTop();
    }

    public static CTBorder getTableTop(CTTbl table) {
        CTTblBorders tblBorders = getTblBorders(table);
        return tblBorders.isSetTop() ? tblBorders.getTop() : tblBorders.addNewTop();
    }

    public static CTBorder getTableCellTop(CTTc cell) {
        CTTcPr tcPr = getTcPr(cell);
        CTTcBorders tcBorders = getTcBorders(tcPr);
        return tcBorders.isSetTop() ? tcBorders.getTop() : tcBorders.addNewTop();
    }

    private static CTBorder getRight(Object e) {
        if (e instanceof XWPFParagraph) {
            XWPFParagraph paragraph = (XWPFParagraph) e;
            return getParagraphRight(paragraph.getCTP());
        } else if (e instanceof XWPFTable) {
            XWPFTable table = (XWPFTable) e;
            return getTableRight(table.getCTTbl());
        } else if (e instanceof XWPFTableCell) {
            XWPFTableCell cell = (XWPFTableCell) e;
            return getTableCellRight(cell.getCTTc());
        } else {
            throw new UnsupportedOperationException("Can not get right border of " + e.getClass().getName());
        }
    }

    public static CTBorder getParagraphRight(CTP paragraph) {
        CTPPr pPr = getPPr(paragraph);
        CTPBdr pBdr = getPBdr(pPr);
        return pBdr.isSetRight() ? pBdr.getRight() : pBdr.addNewRight();
    }

    public static CTBorder getTableRight(CTTbl tbl) {
        CTTblBorders tblBorders = getTblBorders(tbl);
        return tblBorders.isSetRight() ? tblBorders.getRight() : tblBorders.addNewRight();
    }

    public static CTBorder getTableCellRight(CTTc cell) {
        CTTcPr tcPr = getTcPr(cell);
        CTTcBorders tcBorders = getTcBorders(tcPr);
        return tcBorders.isSetRight() ? tcBorders.getRight() : tcBorders.addNewRight();
    }

    private static CTBorder getBottom(Object e) {
        if (e instanceof XWPFParagraph) {
            XWPFParagraph paragraph = (XWPFParagraph) e;
            return getParagraphBottom(paragraph.getCTP());
        } else if (e instanceof XWPFTable) {
            XWPFTable table = (XWPFTable) e;
            return getTableBottom(table.getCTTbl());
        } else if (e instanceof XWPFTableCell) {
            XWPFTableCell cell = (XWPFTableCell) e;
            return getTableCellBottom(cell.getCTTc());
        } else {
            throw new UnsupportedOperationException("Can not get bottom border of " + e.getClass().getName());
        }
    }

    public static CTBorder getParagraphBottom(CTP paragraph) {
        CTPPr pPr = getPPr(paragraph);
        CTPBdr pBdr = getPBdr(pPr);
        return pBdr.isSetBottom() ? pBdr.getBottom() : pBdr.addNewBottom();
    }

    public static CTBorder getTableBottom(CTTbl table) {
        CTTblBorders tblBorders = getTblBorders(table);
        return tblBorders.isSetBottom() ? tblBorders.getBottom() : tblBorders.addNewBottom();
    }

    public static CTBorder getTableCellBottom(CTTc cell) {
        CTTcPr tcPr = getTcPr(cell);
        CTTcBorders tcBorders = getTcBorders(tcPr);
        return tcBorders.isSetBottom() ? tcBorders.getBottom() : tcBorders.addNewBottom();
    }

    private static CTBorder getLeft(Object e) {
        if (e instanceof XWPFParagraph) {
            XWPFParagraph paragraph = (XWPFParagraph) e;
            return getParagraphLeft(paragraph.getCTP());
        } else if (e instanceof XWPFTable) {
            XWPFTable table = (XWPFTable) e;
            return getTableLeft(table.getCTTbl());
        } else if (e instanceof XWPFTableCell) {
            XWPFTableCell cell = (XWPFTableCell) e;
            return getTableCellLeft(cell.getCTTc());
        } else {
            throw new UnsupportedOperationException("Can not get left border of " + e.getClass().getName());
        }
    }

    public static CTBorder getParagraphLeft(CTP paragraph) {
        CTPPr pPr = getPPr(paragraph);
        CTPBdr pBdr = getPBdr(pPr);
        return pBdr.isSetLeft() ? pBdr.getLeft() : pBdr.addNewLeft();
    }

    public static CTBorder getTableLeft(CTTbl table) {
        CTTblBorders tblBorders = getTblBorders(table);
        return tblBorders.isSetLeft() ? tblBorders.getLeft() : tblBorders.addNewLeft();
    }

    public static CTBorder getTableCellLeft(CTTc cell) {
        CTTcPr tcPr = getTcPr(cell);
        CTTcBorders tcBorders = getTcBorders(tcPr);
        return tcBorders.isSetLeft() ? tcBorders.getLeft() : tcBorders.addNewLeft();
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
     * poi 5.x 版本{@link Units}中的该方法被删除了
     *
     * @param twips (1/20th of a point) typically used for row heights
     * @return equivalent EMUs
     */
    public static int twipsToEMU(int twips) {
        return twips * Units.EMU_PER_DXA;
    }

    /**
     * EMU转twip
     *
     * @see #twipsToEMU
     */
    public static int emuToTwips(int emu) {
        return emu / Units.EMU_PER_DXA;
    }

    /**
     * 应用表格样式
     *
     * @param context 渲染上下文
     * @param table 表格
     * @param cssStyleDeclaration CSS样式声明
     */
    public static void tableStyle(HtmlRenderContext context, XWPFTable table, CSSStyleDeclarationImpl cssStyleDeclaration) {
        if (CSSStyleUtils.isEmpty(cssStyleDeclaration)) {
            return;
        }

        // alignment
        TableRowAlign align = alignTable(cssStyleDeclaration.getPropertyValue(HtmlConstants.CSS_FLOAT));
        if (align != null) {
            table.setTableAlignment(align);
        }

        // border
        boolean allNone = setBorder(context, table, cssStyleDeclaration);
        // 如果四边都是none则将单元格间的边框也置为none
        if (allNone) {
            CTTblBorders tblBorders = getTblBorders(table.getCTTbl());
            CTBorder insideH = getTblInsideH(tblBorders);
            insideH.setVal(STBorder.NONE);
            CTBorder insideV = getTblInsideV(tblBorders);
            insideV.setVal(STBorder.NONE);
        }

        // indent
        String marginLeft = cssStyleDeclaration.getPropertyValue(HtmlConstants.CSS_MARGIN_LEFT);
        if (StringUtils.isNotBlank(marginLeft)) {
            indent(context, table, marginLeft);
        }

        // background
        String backgroundColor = cssStyleDeclaration.getBackgroundColor();
        if (StringUtils.isNotBlank(backgroundColor)) {
            String color = Colors.fromStyle(backgroundColor, null);
            if (color != null) {
                CTTblPr tblPr = getTblPr(table.getCTTbl());
                CTShd shd = getShd(tblPr);
                shd.setFill(color);
            }
        }
    }

    public static CTBorder getTblInsideV(CTTblBorders tblBorders) {
        return tblBorders.isSetInsideV() ? tblBorders.getInsideV() : tblBorders.addNewInsideV();
    }

    public static CTBorder getTblInsideH(CTTblBorders tblBorders) {
        return tblBorders.isSetInsideH() ? tblBorders.getInsideH() : tblBorders.addNewInsideH();
    }


    /**
     * 应用表格样式
     *
     * @param context 渲染上下文
     * @param cell 表格
     * @param cssStyleDeclaration CSS样式声明
     */
    public static void cellStyle(HtmlRenderContext context, XWPFTableCell cell, CSSStyleDeclarationImpl cssStyleDeclaration) {
        if (CSSStyleUtils.isEmpty(cssStyleDeclaration)) {
            return;
        }

        // padding
        setCellPadding(context, cell, cssStyleDeclaration);

        // alignment
        XWPFTableCell.XWPFVertAlign align = alignTableCell(cssStyleDeclaration.getVerticalAlign());
        if (align != null) {
            cell.setVerticalAlignment(align);
        }

        // border
        setBorder(context, cell, cssStyleDeclaration);

        // background
        String backgroundColor = cssStyleDeclaration.getBackgroundColor();
        if (StringUtils.isNotBlank(backgroundColor)) {
            String color = Colors.fromStyle(backgroundColor, null);
            if (color != null) {
                CTTcPr tcPr = getTcPr(cell.getCTTc());
                CTShd shd = getShd(tcPr);
                shd.setFill(color);
            }
        }
    }

    /**
     * 设置单元格边距
     *
     * @param context 渲染上下文
     * @param cell 表格
     * @param cssStyleDeclaration CSS样式声明
     */
    private static void setCellPadding(HtmlRenderContext context, XWPFTableCell cell,
                                       CSSStyleDeclarationImpl cssStyleDeclaration) {
        // margin-top
        CSSLength paddingTop = CSSLength.of(cssStyleDeclaration.getPaddingTop().toLowerCase());
        if (paddingTop.isValid() && !paddingTop.isPercent() && paddingTop.getValue() >= 0) {
            CTTblWidth top = getTcMar(cell).addNewTop();
            top.setType(STTblWidth.DXA);
            top.setW(BigInteger.valueOf(emuToTwips(context.lengthToEMU(paddingTop))));
        }

        // margin-right
        CSSLength paddingRight = CSSLength.of(cssStyleDeclaration.getPaddingRight().toLowerCase());
        if (paddingRight.isValid() && !paddingRight.isPercent() && paddingRight.getValue() >= 0) {
            CTTblWidth right = getTcMar(cell).addNewRight();
            right.setType(STTblWidth.DXA);
            right.setW(BigInteger.valueOf(emuToTwips(context.lengthToEMU(paddingRight))));
        }

        // margin-bottom
        CSSLength paddingBottom = CSSLength.of(cssStyleDeclaration.getPaddingBottom().toLowerCase());
        if (paddingBottom.isValid() && !paddingBottom.isPercent() && paddingBottom.getValue() >= 0) {
            CTTblWidth bottom = getTcMar(cell).addNewBottom();
            bottom.setType(STTblWidth.DXA);
            bottom.setW(BigInteger.valueOf(emuToTwips(context.lengthToEMU(paddingBottom))));
        }

        // margin-left
        CSSLength paddingLeft = CSSLength.of(cssStyleDeclaration.getPaddingLeft().toLowerCase());
        if (paddingLeft.isValid() && !paddingLeft.isPercent() && paddingLeft.getValue() >= 0) {
            CTTblWidth left = getTcMar(cell).addNewLeft();
            left.setType(STTblWidth.DXA);
            left.setW(BigInteger.valueOf(emuToTwips(context.lengthToEMU(paddingLeft))));
        }
    }

    /**
     * 设置上下左右边框样式
     *
     * @param xwpfElement 元素
     * @param cssStyleDeclaration CSS样式声明
     * @return 是否四边全部为none
     */
    public static boolean setBorder(HtmlRenderContext context, Object xwpfElement, CSSStyleDeclarationImpl cssStyleDeclaration) {
        boolean topNone = setBorder(context, xwpfElement, cssStyleDeclaration, HtmlConstants.CSS_BORDER_TOP_STYLE,
                HtmlConstants.CSS_BORDER_TOP_WIDTH, HtmlConstants.CSS_BORDER_TOP_COLOR, RenderUtils::getTop);
        boolean rightNone = setBorder(context, xwpfElement, cssStyleDeclaration, HtmlConstants.CSS_BORDER_RIGHT_STYLE,
                HtmlConstants.CSS_BORDER_RIGHT_WIDTH, HtmlConstants.CSS_BORDER_RIGHT_COLOR, RenderUtils::getRight);
        boolean bottomNone = setBorder(context, xwpfElement, cssStyleDeclaration, HtmlConstants.CSS_BORDER_BOTTOM_STYLE,
                HtmlConstants.CSS_BORDER_BOTTOM_WIDTH, HtmlConstants.CSS_BORDER_BOTTOM_COLOR, RenderUtils::getBottom);
        boolean leftNone = setBorder(context, xwpfElement, cssStyleDeclaration, HtmlConstants.CSS_BORDER_LEFT_STYLE,
                HtmlConstants.CSS_BORDER_LEFT_WIDTH, HtmlConstants.CSS_BORDER_LEFT_COLOR, RenderUtils::getLeft);
        return topNone && rightNone && bottomNone && leftNone;
    }

    private static boolean indent(HtmlRenderContext context, XWPFTable table, String style) {
        CSSLength cssLength = CSSLength.of(style.toLowerCase());
        if (cssLength.isValid() && cssLength.getValue() > 0) {
            CTTblPr tblPr = getTblPr(table.getCTTbl());
            CTTblWidth ind = getInd(tblPr);
            double indent;
            if (cssLength.isPercent()) {
                indent = context.getAvailableWidthInEMU() * cssLength.unitValue() / CSSLengthUnit.TWIP.absoluteFactor();
            } else {
                indent = context.lengthToEMU(cssLength) / CSSLengthUnit.TWIP.absoluteFactor();
            }
            ind.setType(STTblWidth.DXA);
            ind.setW(BigInteger.valueOf(Math.round(indent)));
            return true;
        }
        return false;
    }

    public static CTTblWidth getInd(CTTblPr tblPr) {
        return tblPr.isSetTblInd() ? tblPr.getTblInd() : tblPr.addNewTblInd();
    }

    public static CTTblBorders getTblBorders(CTTblPr tblPr) {
        return tblPr.isSetTblBorders() ? tblPr.getTblBorders() : tblPr.addNewTblBorders();
    }

    public static CTShd getShd(CTTblPr tblPr) {
        return tblPr.isSetShd() ? tblPr.getShd() : tblPr.addNewShd();
    }

    public static CTTblPr getTblPr(CTTbl ctTbl) {
        CTTblPr tblPr = ctTbl.getTblPr();
        if (tblPr == null) {
            tblPr = ctTbl.addNewTblPr();
        }
        return tblPr;
    }

    public static CTTcBorders getTcBorders(CTTcPr tcPr) {
        return tcPr.isSetTcBorders() ? tcPr.getTcBorders() : tcPr.addNewTcBorders();
    }

    public static CTShd getShd(CTTcPr tcPr) {
        return tcPr.isSetShd() ? tcPr.getShd() : tcPr.addNewShd();
    }

    /**
     * 表格对齐值映射
     *
     * @param cssFloat 表格对齐样式值
     * @return Word表格对齐枚举
     */
    public static TableRowAlign alignTable(String cssFloat) {
        if (StringUtils.isBlank(cssFloat)) {
            return null;
        }
        switch (cssFloat.toLowerCase()) {
            case HtmlConstants.LEFT:
                return TableRowAlign.LEFT;
            case HtmlConstants.RIGHT:
                return TableRowAlign.RIGHT;
            default:
                return null;
        }
    }

    /**
     * 表格单元格垂直对齐值映射
     *
     * @param verticalAlign 垂直对齐值
     * @return Word表格单元格垂直对齐枚举
     */
    public static XWPFTableCell.XWPFVertAlign alignTableCell(String verticalAlign) {
        if (StringUtils.isBlank(verticalAlign)) {
            return null;
        }
        switch (verticalAlign.toLowerCase()) {
            case HtmlConstants.MIDDLE:
                return XWPFTableCell.XWPFVertAlign.CENTER;
            case HtmlConstants.BOTTOM:
                return XWPFTableCell.XWPFVertAlign.BOTTOM;
            default:
                return XWPFTableCell.XWPFVertAlign.TOP;
        }
    }

    /**
     * 嵌入式图片转换为环绕式图片
     *
     * @param drawing 绘图容器
     * @return 环绕式图片
     */
    public static CTAnchor inlineToAnchor(CTDrawing drawing) {
        CTInline ctInline = drawing.getInlineArray(0);
        CTAnchor ctAnchor = drawing.addNewAnchor();
        ctAnchor.setDocPr(ctInline.getDocPr());
        ctAnchor.setExtent(ctInline.getExtent());
        ctAnchor.setGraphic(ctInline.getGraphic());
        if (ctInline.isSetCNvGraphicFramePr()) {
            ctAnchor.setCNvGraphicFramePr(ctInline.getCNvGraphicFramePr());
        }
        drawing.removeInline(0);

        ctAnchor.setAllowOverlap(false);
        ctAnchor.setBehindDoc(false);
        ctAnchor.setRelativeHeight(0);
        ctAnchor.setDistL(0);
        ctAnchor.setDistR(0);
        ctAnchor.setDistB(0);
        ctAnchor.setDistT(0);
        ctAnchor.setLayoutInCell(true);
        ctAnchor.setLocked(false);

        ctAnchor.setSimplePos2(false);
        CTPoint2D simplePos = ctAnchor.addNewSimplePos();
        simplePos.setX(0);
        simplePos.setY(0);

        return ctAnchor;
    }
}

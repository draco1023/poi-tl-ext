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

package org.ddr.poi.html.tag;

import com.steadystate.css.dom.CSSStyleDeclarationImpl;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.ddr.poi.html.util.CSSLength;
import org.ddr.poi.html.util.CSSLengthUnit;
import org.ddr.poi.html.util.CSSStyleUtils;
import org.ddr.poi.html.util.ColumnStyle;
import org.ddr.poi.html.util.JsoupUtils;
import org.ddr.poi.html.util.RenderUtils;
import org.ddr.poi.html.util.Span;
import org.ddr.poi.html.util.SpanWidth;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.select.Elements;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

/**
 * 表格渲染器
 *
 * @author Draco
 * @since 2021-02-18
 */
public class TableRenderer implements ElementRenderer {
    private static final String[] TAGS = {HtmlConstants.TAG_TABLE};

    private final CSSStyleDeclarationImpl defaultCaptionStyle = new CSSStyleDeclarationImpl();

    public TableRenderer() {
        defaultCaptionStyle.setBackgroundColor("");
        defaultCaptionStyle.setBorder("");
        defaultCaptionStyle.setPadding("");
        defaultCaptionStyle.setMargin("");
        defaultCaptionStyle.setTextAlign(HtmlConstants.CENTER);

        CSSStyleUtils.split(defaultCaptionStyle);
    }

    /**
     * 开始渲染
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @return 是否继续渲染子元素
     */
    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        CSSStyleDeclarationImpl styleDeclaration = context.currentElementStyle();
        String widthDeclaration = styleDeclaration.getWidth();

        XWPFTable table = context.getClosestTable();
        int containerWidth = context.getAvailableWidthInEMU();

        CSSLength width = CSSLength.of(widthDeclaration);
        boolean explicitWidth = width.isValid() && !width.isPercent();
        int tableWidth = context.computeLengthInEMU(widthDeclaration, styleDeclaration.getMaxWidth(), containerWidth, containerWidth);
        int originWidth = !width.isValid() || width.isPercent() ? tableWidth : context.lengthToEMU(width);

        Element caption = JsoupUtils.firstChild(element, HtmlConstants.TAG_CAPTION);
        if (caption != null) {
            renderCaption(context, table, caption);
        }

        Element colgroup = JsoupUtils.firstChild(element, HtmlConstants.TAG_COLGROUP);
        List<ColumnStyle> columnStyles = extractColumnStyles(colgroup);

        Elements trs = JsoupUtils.childRows(element);
        Map<Integer, Span> rowSpanMap = new TreeMap<>();
        TreeMap<Integer, CSSLength> colWidthMap = new TreeMap<>();
        LinkedHashSet<SpanWidth> spanWidths = new LinkedHashSet<>();
        for (int r = 0; r < trs.size(); r++) {
            Element tr = trs.get(r);
            XWPFTableRow row = createRow(table, r);

            Elements tds = JsoupUtils.children(tr, HtmlConstants.TAG_TH, HtmlConstants.TAG_TD);
            int columnIndex = 0;
            int minRowSpan = 1;
            int vMergeCount = 0;
            for (int c = 0; c < tds.size(); c++) {
                Element td = tds.get(c);
                int rowspan = NumberUtils.toInt(td.attr(HtmlConstants.ATTR_ROWSPAN), 1);
                int colspan = NumberUtils.toInt(td.attr(HtmlConstants.ATTR_COLSPAN), 1);
                minRowSpan = Math.min(minRowSpan, rowspan);

                for (Map.Entry<Integer, Span> entry : rowSpanMap.entrySet()) {
                    if (entry.getKey() <= columnIndex && entry.getValue().isEnabled()) {
                        columnIndex += entry.getValue().getColumn();
                        entry.getValue().setEnabled(false);
                        // 合并行也需要生成单元格
                        addVMergeCell(row, c, entry.getValue());
                        vMergeCount++;
                    }
                }
                // 标记行列索引，便于渲染单元格时获取容器
                td.attr(HtmlConstants.ATTR_ROW_INDEX, String.valueOf(r));
                td.attr(HtmlConstants.ATTR_COLUMN_INDEX, String.valueOf(c + vMergeCount));

                // 列定义的样式与单元格的样式合并
                if (!columnStyles.isEmpty() && columnIndex < columnStyles.size()) {
                    String colStyle = columnStyles.get(columnIndex).getStyle().getCssText();
                    StringBuilder sb = new StringBuilder();
                    if (!colStyle.isEmpty()) {
                        sb.append(colStyle).append(HtmlConstants.SEMICOLON);
                    }
                    if (colspan > 1) {
                        CSSLength tdWidth = sumColumnWidths(columnStyles, columnIndex, colspan);
                        if (tdWidth.isValid()) {
                            sb.append(HtmlConstants.CSS_WIDTH).append(HtmlConstants.COLON)
                                    .append(tdWidth).append(HtmlConstants.SEMICOLON);
                        }
                    }
                    if (sb.length() > 0) {
                        sb.append(td.attr(HtmlConstants.ATTR_STYLE));
                        td.attr(HtmlConstants.ATTR_STYLE, sb.toString());
                    }
                }

                CSSStyleDeclarationImpl tdStyleDeclaration = CSSStyleUtils.parseNew(td.attr(HtmlConstants.ATTR_STYLE));
                CSSLength tdWidth = CSSLength.of(tdStyleDeclaration.getWidth());

                // 必须晚于之前列的行合并单元格创建
                XWPFTableCell cell = createCell(row, c);
                CTTcPr ctTcPr = RenderUtils.getTcPr(cell.getCTTc());
                if (rowspan > 1) {
                    CSSStyleUtils.split(tdStyleDeclaration);
                    rowSpanMap.put(columnIndex, new Span(rowspan, colspan, false, tdStyleDeclaration));
                    CTVMerge ctvMerge = ctTcPr.isSetVMerge() ? ctTcPr.getVMerge() : ctTcPr.addNewVMerge();
                    ctvMerge.setVal(STMerge.RESTART);
                }
                if (colspan == 1) {
                    CSSLength existingWidth = colWidthMap.get(columnIndex);
                    if (existingWidth == null || !existingWidth.isValid()) {
                        colWidthMap.put(columnIndex, tdWidth);
                    } else {
                        // 根据表格本身是否使用百分比的宽度定义，来确定单元格宽度的定义方式
                        if (explicitWidth) {
                            if (existingWidth.isPercent()) {
                                colWidthMap.put(columnIndex, tdWidth);
                            }
                        } else {
                            if (!existingWidth.isPercent()) {
                                colWidthMap.put(columnIndex, tdWidth);
                            }
                        }
                    }
                } else {
                    spanWidths.add(new SpanWidth(tdWidth, columnIndex, colspan, explicitWidth));
                    ctTcPr.addNewGridSpan().setVal(BigInteger.valueOf(colspan));
                }

                columnIndex += colspan;
            }

            for (Iterator<Map.Entry<Integer, Span>> iterator = rowSpanMap.entrySet().iterator(); iterator.hasNext(); ) {
                Map.Entry<Integer, Span> entry = iterator.next();
                Integer spanColumnIndex = entry.getKey();
                Span span = entry.getValue();
                span.setRow(span.getRow() - minRowSpan);
                if (span.getRow() == 0) {
                    iterator.remove();
                } else {
                    span.setEnabled(true);
                }
                if (columnIndex <= spanColumnIndex && columnIndex < colWidthMap.size()) {
                    addVMergeCell(row, columnIndex, span);
                    columnIndex += span.getColumn();
                }
            }
        }

        CTTbl ctTbl = table.getCTTbl();
        CTTblGrid tblGrid = ctTbl.getTblGrid();
        if (tblGrid == null) {
            tblGrid = ctTbl.addNewTblGrid();
        }
        for (SpanWidth spanWidth : spanWidths) {
            spanWidth.setLength(colWidthMap);
        }

        BigInteger[] colWidths = new BigInteger[colWidthMap.size()];
        // 未处理的百分比总和
        double unhandledPercentSum = 0;
        // 未处理的emu长度总和
        int unhandledEmuSum = 0;
        // 未处理的emu长度数量
        int unhandledEmuCount = 0;
        // 可用的宽度
        int remainWidth = tableWidth;
        for (Iterator<Map.Entry<Integer, CSSLength>> iterator = colWidthMap.entrySet().iterator(); iterator.hasNext(); ) {
            Map.Entry<Integer, CSSLength> entry = iterator.next();
            CSSLength value = entry.getValue();
            if (!value.isValid()) {
                entry.setValue(new CSSLength(100d / colWidths.length, CSSLengthUnit.PERCENT));
                unhandledPercentSum += entry.getValue().getValue();
                continue;
            }

            if (explicitWidth) {
                if (value.isPercent()) {
                    unhandledPercentSum += value.getValue();
                } else {
                    int emu = (int) ((long) value.toEMU() * tableWidth / originWidth);
                    colWidths[entry.getKey()] = BigInteger.valueOf(RenderUtils.emuToTwips(emu));
                    remainWidth -= emu;
                    iterator.remove();
                }
            } else {
                if (value.isPercent()) {
                    unhandledPercentSum += value.getValue();
                } else {
                    unhandledPercentSum += 100d / colWidths.length;
                    unhandledEmuSum += value.toEMU();
                    unhandledEmuCount++;
                }
            }
        }
        // 处理剩余未处理的列宽
        for (Map.Entry<Integer, CSSLength> entry : colWidthMap.entrySet()) {
            CSSLength value = entry.getValue();
            if (value.isPercent()) {
                colWidths[entry.getKey()] = BigInteger.valueOf((int) Math.rint(remainWidth * value.getValue() / unhandledPercentSum * 20 / Units.EMU_PER_POINT));
            } else {
                colWidths[entry.getKey()] = BigInteger.valueOf((int) Math.rint(remainWidth * (unhandledEmuCount * 100d / colWidths.length / unhandledPercentSum) * value.toEMU() / unhandledEmuSum * 20 / Units.EMU_PER_POINT));
            }
        }
        for (BigInteger colWidth : colWidths) {
            CTTblGridCol ctTblGridCol = tblGrid.addNewGridCol();
            ctTblGridCol.setW(colWidth);
        }
        for (int i = 0, rows = ctTbl.sizeOfTrArray(); i < rows; i++) {
            CTRow ctRow = ctTbl.getTrArray(i);
            int columnIndex = 0;
            for (int j = 0, cells = ctRow.sizeOfTcArray(); j < cells; j++) {
                CTTc ctTc = ctRow.getTcArray(j);
                CTTcPr tcPr = RenderUtils.getTcPr(ctTc);
                int colspan = tcPr.isSetGridSpan() ? tcPr.getGridSpan().getVal().intValue() : 1;
                CTTblWidth tcWidth = tcPr.addNewTcW();
                tcWidth.setType(STTblWidth.DXA);
                if (colspan == 1) {
                    tcWidth.setW(colWidths[columnIndex]);
                } else {
                    int sum = 0;
                    for (int k = 0; k < colspan; k++) {
                        sum += colWidths[columnIndex + k].intValue();
                    }
                    tcWidth.setW(BigInteger.valueOf(sum));
                }
                columnIndex += colspan;
            }
        }

        return true;
    }

    private List<ColumnStyle> extractColumnStyles(Element colgroup) {
        List<ColumnStyle> columnStyles = Collections.emptyList();
        if (colgroup != null) {
            Elements cols = colgroup.select(HtmlConstants.TAG_COL);
            columnStyles = new ArrayList<>();
            for (Element col : cols) {
                String style = col.attr(HtmlConstants.ATTR_STYLE);
                CSSStyleDeclarationImpl cssStyleDeclaration = CSSStyleUtils.parse(style);
                int span = NumberUtils.toInt(col.attr(HtmlConstants.ATTR_SPAN), 1);
                // 宽度样式优先于宽度属性
                CSSLength colWidth = CSSLength.of(cssStyleDeclaration.getWidth());
                if (!colWidth.isValid()) {
                    String colWidthAttr = col.attr(HtmlConstants.ATTR_WIDTH);
                    if (!colWidthAttr.isEmpty()) {
                        if (!colWidthAttr.endsWith(HtmlConstants.PERCENT)) {
                            colWidthAttr += HtmlConstants.PX;
                        }
                        colWidth = CSSLength.of(colWidthAttr);
                    }
                }
                if (colWidth.isValid() && span > 1) {
                    colWidth = new CSSLength(colWidth.getValue() / span, colWidth.getUnit());
                }
                for (int i = 0; i < span; i++) {
                    columnStyles.add(new ColumnStyle(cssStyleDeclaration, colWidth));
                }
            }
            colgroup.remove();
        }
        return columnStyles;
    }

    private CSSLength sumColumnWidths(List<ColumnStyle> columnWidths, int columnIndex, int colspan) {
        Boolean percent = null;
        double sum = 0d;
        for (int i = 0; i < colspan; i++) {
            CSSLength width = columnWidths.get(columnIndex + i).getWidth();
            if (!width.isValid()) {
                return width;
            }
            if (percent == null) {
                percent = width.isPercent();
            } else if (percent ^ width.isPercent()) {
                return CSSLength.INVALID;
            }
            if (percent) {
                sum += width.getValue();
            } else {
                sum += width.toEMU();
            }
        }
        return new CSSLength(sum, percent ? CSSLengthUnit.PERCENT : CSSLengthUnit.EMU);
    }

    /**
     * 渲染标题
     *
     * @param context 渲染上下文
     * @param table 表格
     * @param caption 标题元素
     */
    private void renderCaption(HtmlRenderContext context, XWPFTable table, Element caption) {
        CSSStyleDeclarationImpl captionStyle = context.getCssStyleDeclaration(caption);
        if (CSSStyleUtils.EMPTY_STYLE == captionStyle) {
            captionStyle = defaultCaptionStyle;
        } else {
            captionStyle.getProperties().addAll(0, defaultCaptionStyle.getProperties());
        }
        context.pushInlineStyle(captionStyle, caption.isBlock());
        XWPFParagraph captionParagraph;
        boolean bottom = HtmlConstants.BOTTOM.equals(context.getPropertyValue(HtmlConstants.CSS_CAPTION_SIDE));
        if (bottom) {
            captionParagraph = context.getClosestParagraph();
            caption.parent().attr(HtmlConstants.CSS_CAPTION_SIDE, HtmlConstants.BOTTOM);
        } else {
            // 表格上方添加标题
            XmlCursor xmlCursor = table.getCTTbl().newCursor();
            context.pushCursor(xmlCursor);
            captionParagraph = context.newParagraph(null, xmlCursor);
            xmlCursor.dispose();
        }
        RenderUtils.paragraphStyle(context, captionParagraph, captionStyle);

        context.markDedupe(captionParagraph);
        for (Node node : caption.childNodes()) {
            context.renderNode(node);
        }
        context.unmarkDedupe();

        context.popInlineStyle();
        caption.remove();
        if (bottom) {
            XmlCursor xmlCursor = captionParagraph.getCTP().newCursor();
            context.pushCursor(xmlCursor);
            xmlCursor.dispose();
        } else {
            context.popCursor();
        }
    }

    /**
     * 元素渲染结束需要执行的逻辑
     *
     * @param element HTML元素
     * @param context 渲染上下文
     */
    @Override
    public void renderEnd(Element element, HtmlRenderContext context) {
        if (HtmlConstants.BOTTOM.equals(element.attr(HtmlConstants.CSS_CAPTION_SIDE))) {
            context.popCursor();
        }
    }

    /**
     * 添加垂直合并的单元格
     *
     * @param row 行
     * @param columnIndex 列索引
     * @param span 跨行列参数
     */
    private void addVMergeCell(XWPFTableRow row, int columnIndex, Span span) {
        XWPFTableCell cell = createCell(row, columnIndex);
        CTTcPr ctTcPr = RenderUtils.getTcPr(cell.getCTTc());
        ctTcPr.addNewVMerge();
        RenderUtils.setBorder(cell, span.getStyle());
        if (span.getColumn() > 1) {
            ctTcPr.addNewGridSpan().setVal(BigInteger.valueOf(span.getColumn()));
        }
    }

    /**
     * 创建单元格
     *
     * @param row 行
     * @param c 列索引
     * @return 单元格
     */
    private XWPFTableCell createCell(XWPFTableRow row, int c) {
        return row.createCell();
    }

    /**
     * 创建行
     *
     * @param table 表格
     * @param r 行索引
     * @return 行
     */
    private XWPFTableRow createRow(XWPFTable table, int r) {
        // 避免使用createRow，因为不需要自动创建单元格
        return table.insertNewTableRow(r);
    }

    @Override
    public String[] supportedTags() {
        return TAGS;
    }

    @Override
    public boolean renderAsBlock() {
        return true;
    }
}

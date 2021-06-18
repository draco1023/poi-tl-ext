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
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.ddr.poi.html.util.CSSLength;
import org.ddr.poi.html.util.CSSLengthUnit;
import org.ddr.poi.html.util.JsoupUtils;
import org.ddr.poi.html.util.RenderUtils;
import org.ddr.poi.html.util.Span;
import org.jsoup.nodes.Element;
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
import java.util.HashMap;
import java.util.Iterator;
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

        // FIXME poi不支持设置tblCaption
//        Element caption = JsoupUtils.firstChild(element, HtmlConstants.TAG_CAPTION);

        Element colgroup = JsoupUtils.firstChild(element, HtmlConstants.TAG_COLGROUP);
        if (colgroup != null) {
            Elements cols = colgroup.select(HtmlConstants.TAG_COL);
            // FIXME col可以拥有style
            colgroup.remove();
        }
        Elements trs = JsoupUtils.childRows(element);
        Map<Integer, Span> rowSpanMap = new HashMap<>(4);
        TreeMap<Integer, CSSLength> colWidthMap = new TreeMap<>();
        for (int r = 0; r < trs.size(); r++) {
            Element tr = trs.get(r);
            XWPFTableRow row = createRow(table, r);

            Elements tds = JsoupUtils.children(tr, HtmlConstants.TAG_TH, HtmlConstants.TAG_TD);
            int columnSum = 0;
            int minRowSpan = 1;
            for (int c = 0; c < tds.size(); c++) {
                Element td = tds.get(c);
                CSSStyleDeclarationImpl tdStyleDeclaration = RenderUtils.parse(td.attr(HtmlConstants.ATTR_STYLE));
                CSSLength tdWidth = CSSLength.of(tdStyleDeclaration.getWidth());
                int rowspan = NumberUtils.toInt(td.attr(HtmlConstants.ATTR_ROWSPAN), 1);
                int colspan = NumberUtils.toInt(td.attr(HtmlConstants.ATTR_COLSPAN), 1);
                minRowSpan = Math.min(minRowSpan, rowspan);
                for (Map.Entry<Integer, Span> entry : rowSpanMap.entrySet()) {
                    if (entry.getKey() <= columnSum && entry.getValue().isEnabled()) {
                        columnSum += entry.getValue().getColumn();
                        entry.getValue().setEnabled(false);
                        // 合并行也需要生成单元格
                        XWPFTableCell cell = createCell(row, c);
                        CTTcPr ctTcPr = RenderUtils.getTcPr(cell.getCTTc());
                        ctTcPr.addNewVMerge();
                        if (entry.getValue().getColumn() > 1) {
                            ctTcPr.addNewGridSpan().setVal(BigInteger.valueOf(entry.getValue().getColumn()));
                        }
                    }
                }
                // 标记行列索引，便于渲染单元格时获取容器
                td.attr(HtmlConstants.ATTR_ROW_INDEX, String.valueOf(r));
                td.attr(HtmlConstants.ATTR_COLUMN_INDEX, String.valueOf(c));

                // 必须晚于之前列的行合并单元格创建
                XWPFTableCell cell = createCell(row, c);
                CTTcPr ctTcPr = RenderUtils.getTcPr(cell.getCTTc());
                if (rowspan > 1) {
                    rowSpanMap.put(columnSum, new Span(rowspan, colspan, false));
                    CTVMerge ctvMerge = ctTcPr.isSetVMerge() ? ctTcPr.getVMerge() : ctTcPr.addNewVMerge();
                    ctvMerge.setVal(STMerge.RESTART);
                }
                if (colspan == 1) {
                    CSSLength existingWidth = colWidthMap.get(columnSum);
                    if (existingWidth == null || !existingWidth.isValid()) {
                        colWidthMap.put(columnSum, tdWidth);
                    } else {
                        // 根据表格本身是否使用百分比的宽度定义，来确定单元格宽度的定义方式
                        if (explicitWidth) {
                            if (existingWidth.isPercent()) {
                                colWidthMap.put(columnSum, tdWidth);
                            }
                        } else {
                            if (!existingWidth.isPercent()) {
                                colWidthMap.put(columnSum, tdWidth);
                            }
                        }
                    }
                } else {
                    ctTcPr.addNewGridSpan().setVal(BigInteger.valueOf(colspan));
                }

                columnSum += colspan;
            }

            for (Iterator<Map.Entry<Integer, Span>> iterator = rowSpanMap.entrySet().iterator(); iterator.hasNext(); ) {
                Map.Entry<Integer, Span> entry = iterator.next();
                entry.getValue().setRow(entry.getValue().getRow() - minRowSpan);
                if (entry.getValue().getRow() == 0) {
                    iterator.remove();
                } else {
                    entry.getValue().setEnabled(true);
                }
            }
        }

        CTTbl ctTbl = table.getCTTbl();
        CTTblGrid tblGrid = ctTbl.getTblGrid();
        if (tblGrid == null) {
            tblGrid = ctTbl.addNewTblGrid();
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
            for (int j = 0, cells = ctRow.sizeOfTcArray(); j < cells; j++) {
                CTTc ctTc = ctRow.getTcArray(j);
                CTTcPr tcPr = RenderUtils.getTcPr(ctTc);
                int colspan = tcPr.isSetGridSpan() ? tcPr.getGridSpan().getVal().intValue() : 1;
                CTTblWidth tcWidth = tcPr.addNewTcW();
                tcWidth.setType(STTblWidth.DXA);
                if (colspan == 1) {
                    tcWidth.setW(colWidths[j]);
                } else {
                    int sum = 0;
                    for (int k = 0; k < colspan; k++) {
                        sum += colWidths[j + k].intValue();
                    }
                    tcWidth.setW(BigInteger.valueOf(sum));
                }
            }
        }

        return true;
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

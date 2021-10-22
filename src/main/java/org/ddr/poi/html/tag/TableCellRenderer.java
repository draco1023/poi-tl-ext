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
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.xmlbeans.XmlCursor;
import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.ddr.poi.html.util.RenderUtils;
import org.jsoup.nodes.Element;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;

import java.util.List;

/**
 * 表格单元格标签渲染器
 *
 * @author Draco
 * @since 2021-03-04
 */
public class TableCellRenderer implements ElementRenderer {
    private static final String[] TAGS = {HtmlConstants.TAG_TH, HtmlConstants.TAG_TD};

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
        int row = NumberUtils.toInt(element.attr(HtmlConstants.ATTR_ROW_INDEX));
        int column = NumberUtils.toInt(element.attr(HtmlConstants.ATTR_COLUMN_INDEX));
        XWPFTable table = context.getClosestTable();
        XWPFTableCell cell = table.getRow(row).getCell(column);
        context.pushContainer(cell);
        context.pushClosestBody(cell.getParagraphArray(0));

        RenderUtils.cellStyle(context, cell, styleDeclaration);

        return true;
    }

    /**
     * 元素渲染结束需要执行的逻辑
     *
     * @param element HTML元素
     * @param context 渲染上下文
     */
    @Override
    public void renderEnd(Element element, HtmlRenderContext context) {
        // 单元格创建时默认包含一个空的段落，如果渲染完成时该段落不包含内容则删除
        List<XWPFParagraph> paragraphs = context.getContainer().getParagraphs();
        if (paragraphs.size() > 1) {
            XmlCursor xmlCursor = paragraphs.get(0).getCTP().newCursor();
            if (!xmlCursor.toFirstChild()) {
                xmlCursor.removeXml();
                paragraphs.remove(0);
            }
            xmlCursor.dispose();
        }

        ParagraphAlignment alignment = RenderUtils.align(context.currentElementStyle().getTextAlign());
        if (alignment != null) {
            for (XWPFParagraph paragraph : paragraphs) {
                CTPPr pPr = paragraph.getCTP().getPPr();
                if (pPr == null || !pPr.isSetJc()) {
                    paragraph.setAlignment(alignment);
                }
            }
        }

        context.popContainer();
        context.popClosestBody();
    }

    @Override
    public String[] supportedTags() {
        return TAGS;
    }

    @Override
    public boolean renderAsBlock() {
        // 本身仅作为容器
        return false;
    }
}

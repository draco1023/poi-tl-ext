package org.ddr.poi.html.tag;

import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.jsoup.nodes.Element;

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
        int row = NumberUtils.toInt(element.attr(HtmlConstants.ATTR_ROW_INDEX));
        int column = NumberUtils.toInt(element.attr(HtmlConstants.ATTR_COLUMN_INDEX));
        XWPFTable table = context.getClosestTable();
        XWPFTableCell cell = table.getRow(row).getCell(column);
        context.pushContainer(cell);
        context.pushClosestBody(cell.getParagraphArray(0));

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

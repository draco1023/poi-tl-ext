package org.ddr.poi.html.tag;

import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.ddr.poi.html.util.RenderUtils;
import org.jsoup.nodes.Element;

/**
 * h1~h6标签渲染器
 *
 * @author Draco
 * @since 2021-02-24
 */
public class HeaderRenderer implements ElementRenderer {
    private static final String[] TAGS = {
            HtmlConstants.TAG_H1, HtmlConstants.TAG_H2, HtmlConstants.TAG_H3,
            HtmlConstants.TAG_H4, HtmlConstants.TAG_H5, HtmlConstants.TAG_H6
    };

    /**
     * 各级别标题对应字号
     */
    private static final String[] FONT_SIZES = {"24pt", "18pt", "14pt", "12pt", "10pt", "7.5pt"};

    /**
     * 开始渲染
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @return 是否继续渲染子元素
     */
    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        int index = Integer.parseInt(element.normalName().substring(1)) - 1;
        String fontSizeStyle = HtmlConstants.inlineStyle(HtmlConstants.CSS_FONT_SIZE, FONT_SIZES[index]);
        context.pushInlineStyle(RenderUtils.parse(HtmlConstants.DEFINED_BOLD + fontSizeStyle), element.isBlock());
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
        context.popInlineStyle();
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

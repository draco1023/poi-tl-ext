package org.ddr.poi.html.tag;

import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.jsoup.nodes.Element;

/**
 * figcaption标签渲染器
 *
 * @author Draco
 * @since 2022-11-03
 */
public class FigureCaptionRenderer implements ElementRenderer {
    private static final String[] TAGS = {HtmlConstants.TAG_FIGURE_CAPTION};

    /**
     * 开始渲染
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @return 是否继续渲染子元素
     */
    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        context.markDedupe(context.getClosestParagraph());
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
        context.unmarkDedupe();
    }

    /**
     * @return 支持的HTML标签
     */
    @Override
    public String[] supportedTags() {
        return TAGS;
    }

    /**
     * @return 是否为块状渲染，如果为true在Word中会另起一个Paragraph
     */
    @Override
    public boolean renderAsBlock() {
        return true;
    }
}

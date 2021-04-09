package org.ddr.poi.html.tag;

import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.jsoup.nodes.Element;

/**
 * 穿透的元素，即标签仅作为样式的载体，但是不渲染为内容容器
 *
 * @author Draco
 * @since 2021-03-15
 */
public class WalkThroughRenderer implements ElementRenderer {
    private static final String[] TAGS = {
            HtmlConstants.TAG_HTML,
            HtmlConstants.TAG_BODY,
            HtmlConstants.TAG_THEAD,
            HtmlConstants.TAG_TBODY,
            HtmlConstants.TAG_TR,
            HtmlConstants.TAG_TFOOT
    };

    /**
     * 开始渲染
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @return 是否继续渲染子元素
     */
    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        return true;
    }

    @Override
    public String[] supportedTags() {
        return TAGS;
    }

    @Override
    public boolean renderAsBlock() {
        return false;
    }
}

package org.ddr.poi.html.tag;

import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.ddr.poi.html.util.CSSStyleUtils;
import org.jsoup.nodes.Element;

/**
 * pre标签渲染器
 *
 * @author Draco
 * @since 2023-06-25
 */
public class PreRenderer implements ElementRenderer {
    private static final String[] TAGS = {HtmlConstants.TAG_PRE, HtmlConstants.TAG_XMP};

    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        context.pushInlineStyle(CSSStyleUtils.parse(HtmlConstants.DEFINED_PRE), element.isBlock());
        return true;
    }

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

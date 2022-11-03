package org.ddr.poi.html.tag;

import com.steadystate.css.dom.CSSStyleDeclarationImpl;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.ddr.poi.html.util.JsoupUtils;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

/**
 * figure标签渲染器
 *
 * @author Draco
 * @since 2022-11-03
 */
public class FigureRenderer implements ElementRenderer {
    private static final String[] TAGS = {HtmlConstants.TAG_FIGURE};

    /**
     * 开始渲染
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @return 是否继续渲染子元素
     */
    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        // https://developer.mozilla.org/en-US/docs/Web/HTML/Element/figure#usage_notes
        Elements captions = JsoupUtils.children(element, HtmlConstants.TAG_FIGURE_CAPTION);
        if (captions.size() > 1) {
            captions.remove(0);
            captions.remove();
        }

        XWPFParagraph paragraph = context.getClosestParagraph();
        context.markDedupe(paragraph);

        CSSStyleDeclarationImpl styleDeclaration = context.currentElementStyle();
        String cssFloat = styleDeclaration.getPropertyValue(HtmlConstants.CSS_FLOAT);
        if (HtmlConstants.LEFT.equals(cssFloat)) {
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            styleDeclaration.setTextAlign(HtmlConstants.LEFT);
        } else if (HtmlConstants.RIGHT.equals(cssFloat)) {
            paragraph.setAlignment(ParagraphAlignment.RIGHT);
            styleDeclaration.setTextAlign(HtmlConstants.RIGHT);
        } else {
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            styleDeclaration.setTextAlign(HtmlConstants.CENTER);
        }

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

package org.ddr.poi.html;

import org.jsoup.nodes.Element;


/**
 * HTML元素渲染器
 *
 * @author Draco
 * @since 2021-02-08
 */
public interface ElementRenderer {
    /**
     * 开始渲染
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @return 是否继续渲染子元素
     */
    boolean renderStart(Element element, HtmlRenderContext context);

    /**
     * 元素渲染结束需要执行的逻辑
     *
     * @param element HTML元素
     * @param context 渲染上下文
     */
    default void renderEnd(Element element, HtmlRenderContext context) {
    }

    /**
     * @return 支持的HTML标签
     */
    String[] supportedTags();

    /**
     * @return 是否为块状渲染，如果为true在Word中会另起一个Paragraph
     */
    boolean renderAsBlock();
}

package org.jsoup.parser;

import org.ddr.poi.html.HtmlConstants;
import org.jsoup.nodes.Element;

/**
 * 自定义html树构建器，对于svg及其内部的标签使用xml解析模式
 *
 * @author Draco
 * @since 2022-04-15
 */
public class CustomHtmlTreeBuilder extends HtmlTreeBuilder {
    @Override
    void reconstructFormattingElements() {
        boolean settingsChanged = false;
        ParseSettings origin = settings;
        if (isSvgElement()) {
            settings = ParseSettings.preserveCase;
            settingsChanged = true;
        }
        super.reconstructFormattingElements();
        if (settingsChanged) {
            settings = origin;
        }
    }

    @Override
    Element insertElementFor(Token.StartTag startTag) {
        boolean settingsChanged = false;
        ParseSettings origin = settings;
        if (isSvgElement()) {
            settings = ParseSettings.preserveCase;
            settingsChanged = true;
        }
        Element element = super.insertElementFor(startTag);
        if (settingsChanged) {
            settings = origin;
        }
        return element;
    }

    private boolean isSvgElement() {
        if (currentToken.isStartTag() && HtmlConstants.TAG_SVG.equals(currentToken.asStartTag().normalName)) {
            return true;
        }
        return getFromStack(HtmlConstants.TAG_SVG) != null;
    }
}

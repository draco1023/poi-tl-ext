package org.jsoup.parser;

import org.ddr.poi.html.HtmlConstants;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;

import java.util.List;

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
    Element insert(Token.StartTag startTag) {
        boolean settingsChanged = false;
        ParseSettings origin = settings;
        if (isSvgElement()) {
            settings = ParseSettings.preserveCase;
            settingsChanged = true;
        }
        Element element = super.insert(startTag);
        if (settingsChanged) {
            settings = origin;
        }
        return element;
    }

    @Override
    public List<Node> parseFragment(String inputFragment, Element context, String baseUri, Parser parser) {
        return super.parseFragment(inputFragment, context, baseUri, parser);
    }

    private boolean isSvgElement() {
        if (currentToken.isStartTag() && HtmlConstants.TAG_SVG.equals(currentToken.asStartTag().normalName)) {
            return true;
        }
        return getFromStack(HtmlConstants.TAG_SVG) != null;
    }
}

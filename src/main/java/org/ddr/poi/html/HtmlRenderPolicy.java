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

package org.ddr.poi.html;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.steadystate.css.dom.CSSStyleDeclarationImpl;
import com.steadystate.css.dom.CSSValueImpl;
import com.steadystate.css.dom.Property;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.BodyType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.xmlbeans.XmlCursor;
import org.ddr.poi.html.tag.ARenderer;
import org.ddr.poi.html.tag.BigRenderer;
import org.ddr.poi.html.tag.BoldRenderer;
import org.ddr.poi.html.tag.BreakRenderer;
import org.ddr.poi.html.tag.DeleteRenderer;
import org.ddr.poi.html.tag.HeaderBreakRenderer;
import org.ddr.poi.html.tag.HeaderRenderer;
import org.ddr.poi.html.tag.ImageRenderer;
import org.ddr.poi.html.tag.ItalicRenderer;
import org.ddr.poi.html.tag.ListItemRenderer;
import org.ddr.poi.html.tag.ListRenderer;
import org.ddr.poi.html.tag.MathRenderer;
import org.ddr.poi.html.tag.OmittedRenderer;
import org.ddr.poi.html.tag.SmallRenderer;
import org.ddr.poi.html.tag.SubscriptRenderer;
import org.ddr.poi.html.tag.SuperscriptRenderer;
import org.ddr.poi.html.tag.TableCellRenderer;
import org.ddr.poi.html.tag.TableRenderer;
import org.ddr.poi.html.tag.UnderlineRenderer;
import org.ddr.poi.html.tag.WalkThroughRenderer;
import org.ddr.poi.html.util.BoxProperty;
import org.ddr.poi.html.util.CSSLength;
import org.ddr.poi.html.util.Colors;
import org.ddr.poi.html.util.NamedBorderWidth;
import org.ddr.poi.html.util.RenderUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.css.sac.InputSource;
import org.w3c.dom.css.CSSValue;

import java.io.IOException;
import java.io.StringReader;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * HTML字符串渲染策略
 *
 * @author Draco
 * @since 2021-02-07
 */
public class HtmlRenderPolicy extends AbstractRenderPolicy<String> {
    private static final Logger log = LoggerFactory.getLogger(HtmlRenderPolicy.class);

    private final Map<String, ElementRenderer> elRenderers;
    private static final Pattern FORMATTED_PATTERN = Pattern.compile(">\\s+<");
    private static final String FORMATTED_REPLACEMENT = "><";
    private final String globalFont;
    private final int globalFontSizeInHalfPoints;

    public HtmlRenderPolicy() {
        this(null, null);
    }

    public HtmlRenderPolicy(String globalFont, CSSLength globalFontSize) {
        ElementRenderer[] renderers = {
                new ARenderer(),
                new BigRenderer(),
                new BoldRenderer(),
                new BreakRenderer(),
                new DeleteRenderer(),
                new HeaderBreakRenderer(),
                new HeaderRenderer(),
                new ImageRenderer(),
                new ItalicRenderer(),
                new ListItemRenderer(),
                new ListRenderer(),
                new MathRenderer(),
                new OmittedRenderer(),
                new SmallRenderer(),
                new SubscriptRenderer(),
                new SuperscriptRenderer(),
                new TableCellRenderer(),
                new TableRenderer(),
                new UnderlineRenderer(),
                new WalkThroughRenderer()
        };
        elRenderers = new HashMap<>(renderers.length);
        for (ElementRenderer renderer : renderers) {
            for (String tag : renderer.supportedTags()) {
                elRenderers.put(tag, renderer);
            }
        }
        this.globalFont = globalFont;
        globalFontSizeInHalfPoints = globalFontSize == null ? 0 : globalFontSize.toHalfPoints();
    }

    @Override
    protected boolean validate(String data) {
        return StringUtils.isNotEmpty(data);
    }

    @Override
    public void doRender(RenderContext<String> context) throws Exception {
        String html = FORMATTED_PATTERN.matcher(context.getData()).replaceAll(FORMATTED_REPLACEMENT);
        Document document = Jsoup.parseBodyFragment(html);

        HtmlRenderContext htmlRenderContext = new HtmlRenderContext(context);
        htmlRenderContext.setGlobalFont(globalFont);
        if (globalFontSizeInHalfPoints > 0) {
            htmlRenderContext.setGlobalFontSize(BigInteger.valueOf(globalFontSizeInHalfPoints));
        }

        for (Node node : document.body().childNodes()) {
            renderNode(node, htmlRenderContext);
        }
    }

    private void renderNode(Node node, HtmlRenderContext context) {
        boolean isElement = node instanceof Element;

        if (isElement) {
            Element element = ((Element) node);
            renderElement(element, context);
        } else if (node instanceof TextNode) {
            context.renderText(((TextNode) node).text());
        }
    }

    private void renderElement(Element element, HtmlRenderContext context) {
        if (log.isDebugEnabled()) {
            log.info("Start rendering html tag: <{}{}>", element.normalName(), element.attributes());
        }
        if (element.tag().isFormListed()) {
            return;
        }

        CSSStyleDeclarationImpl cssStyleDeclaration = getCssStyleDeclaration(element);
        String display = cssStyleDeclaration.getPropertyValue(HtmlConstants.CSS_DISPLAY);
        if (HtmlConstants.NONE.equalsIgnoreCase(display)) {
            return;
        }
        context.pushInlineStyle(cssStyleDeclaration, element.isBlock());

        ElementRenderer elementRenderer = elRenderers.get(element.normalName());

        if (element.isBlock() && (elementRenderer == null || elementRenderer.renderAsBlock())) {
            IBodyElement closestBody = context.getClosestBody();
            XmlCursor xmlCursor;
            switch (closestBody.getElementType()) {
                case PARAGRAPH:
                    xmlCursor = ((XWPFParagraph) closestBody).getCTP().newCursor();
                    break;
                case TABLE:
                    xmlCursor = ((XWPFTable) closestBody).getCTTbl().newCursor();
                    break;
                default:
                    return;
            }

            IBody container = context.getContainer();
            boolean isTableTag = HtmlConstants.TAG_TABLE.equals(element.normalName());
            // 如果是表格，检查当前word容器的前一个兄弟元素是否为表格，是则插入一个段落，防止表格粘连在一起
            if (isTableTag && xmlCursor.toPrevSibling()) {
                if (xmlCursor.getObject() instanceof CTTbl) {
                    xmlCursor.toEndToken();
                    xmlCursor.toNextToken();
                    container.insertNewParagraph(xmlCursor);
                }
                xmlCursor.toNextSibling();
            }

            xmlCursor.toEndToken();
            xmlCursor.toNextToken();

            if (isTableTag) {
                XWPFTable xwpfTable = container.insertNewTbl(xmlCursor);
                xmlCursor.dispose();
                // 新增时会自动创建一行一列，会影响自定义的表格渲染逻辑，故删除
                xwpfTable.removeRow(0);
                context.replaceClosestBody(xwpfTable);

                RenderUtils.tableStyle(context, xwpfTable, cssStyleDeclaration);
            } else {
                XWPFParagraph xwpfParagraph = container.insertNewParagraph(xmlCursor);
                xmlCursor.dispose();
                if (xwpfParagraph == null) {
                    log.warn("Can not add new paragraph for element: {}, attributes: {}", element.tagName(), element.attributes().html());
                }
                context.replaceClosestBody(xwpfParagraph);

                RenderUtils.paragraphStyle(context, xwpfParagraph, cssStyleDeclaration);
            }
        }

        if (elementRenderer != null) {
            if (!elementRenderer.renderStart(element, context)) {
                elementRenderer.renderEnd(element, context);
                context.popInlineStyle();
                return;
            }
        }

        for (Node child : element.childNodes()) {
            renderNode(child, context);
        }

        if (elementRenderer != null) {
            elementRenderer.renderEnd(element, context);
        }
        context.popInlineStyle();
    }

    @Override
    protected void afterRender(RenderContext<String> context) {
        clearPlaceholder(context, true);
        IBody container = context.getContainer();
        if (container.getPartType() == BodyType.TABLECELL) {
            // 单元格的最后一个元素应为p，否则可能无法正常打开文件
            List<IBodyElement> bodyElements = container.getBodyElements();
            if (bodyElements.isEmpty() || bodyElements.get(bodyElements.size() - 1).getElementType() != BodyElementType.PARAGRAPH) {
                ((XWPFTableCell) container).addParagraph();
            }
        }
    }

    private CSSStyleDeclarationImpl getCssStyleDeclaration(Element element) {
        String style = element.attr(HtmlConstants.ATTR_STYLE);
        try (StringReader sr = new StringReader(style)) {
            CSSStyleDeclarationImpl cssStyleDeclaration = (CSSStyleDeclarationImpl) RenderUtils.CSS_PARSER.parseStyleDeclaration(new InputSource(sr));
            for (int i = cssStyleDeclaration.getProperties().size() - 1; i >= 0; i--) {
                final Property p = cssStyleDeclaration.getProperties().get(i);
                if (p != null && p.getValue() != null) {
                    String name = p.getName().toLowerCase();
                    CSSValueImpl valueList = (CSSValueImpl) p.getValue();
                    int length = valueList.getLength();
                    // 将复合样式拆分成单属性样式
                    switch (name) {
                        case HtmlConstants.CSS_BACKGROUND:
                            splitBackground(valueList, length, cssStyleDeclaration, i);
                            break;
                        case HtmlConstants.CSS_BORDER:
                            splitBorder(valueList, length, cssStyleDeclaration, i);
                            break;
                        case HtmlConstants.CSS_BORDER_TOP:
                            splitBorder(valueList, length, cssStyleDeclaration, i, HtmlConstants.CSS_BORDER_TOP_STYLE,
                                    HtmlConstants.CSS_BORDER_TOP_WIDTH, HtmlConstants.CSS_BORDER_TOP_COLOR);
                            break;
                        case HtmlConstants.CSS_BORDER_RIGHT:
                            splitBorder(valueList, length, cssStyleDeclaration, i, HtmlConstants.CSS_BORDER_RIGHT_STYLE,
                                    HtmlConstants.CSS_BORDER_RIGHT_WIDTH, HtmlConstants.CSS_BORDER_RIGHT_COLOR);
                            break;
                        case HtmlConstants.CSS_BORDER_BOTTOM:
                            splitBorder(valueList, length, cssStyleDeclaration, i, HtmlConstants.CSS_BORDER_BOTTOM_STYLE,
                                    HtmlConstants.CSS_BORDER_BOTTOM_WIDTH, HtmlConstants.CSS_BORDER_BOTTOM_COLOR);
                            break;
                        case HtmlConstants.CSS_BORDER_LEFT:
                            splitBorder(valueList, length, cssStyleDeclaration, i, HtmlConstants.CSS_BORDER_LEFT_STYLE,
                                    HtmlConstants.CSS_BORDER_LEFT_WIDTH, HtmlConstants.CSS_BORDER_LEFT_COLOR);
                            break;
                        case HtmlConstants.CSS_BORDER_STYLE:
                            splitBox(valueList, length, cssStyleDeclaration, i, BoxProperty.BORDER_STYLE);
                            break;
                        case HtmlConstants.CSS_BORDER_WIDTH:
                            splitBox(valueList, length, cssStyleDeclaration, i, BoxProperty.BORDER_WIDTH);
                            break;
                        case HtmlConstants.CSS_BORDER_COLOR:
                            splitBox(valueList, length, cssStyleDeclaration, i, BoxProperty.BORDER_COLOR);
                            break;
                        case HtmlConstants.CSS_FONT:
                            splitFont(valueList, length, cssStyleDeclaration, i);
                            break;
                        case HtmlConstants.CSS_MARGIN:
                            splitBox(valueList, length, cssStyleDeclaration, i, BoxProperty.MARGIN);
                            break;
                        case HtmlConstants.CSS_PADDING:
                            splitBox(valueList, length, cssStyleDeclaration, i, BoxProperty.PADDING);
                            break;
                    }
                }
            }
            return cssStyleDeclaration;
        } catch (IOException e) {
            log.warn("Inline style parse error: {}", style, e);
            return RenderUtils.EMPTY_STYLE;
        }
    }

    private void splitBackground(CSSValueImpl valueList, int length, CSSStyleDeclarationImpl cssStyleDeclaration, int i) {
        if (length == 0) {
            String cssText = valueList.getCssText().toLowerCase();
            String color = Colors.fromStyle(cssText, null);
            if (color != null) {
                cssStyleDeclaration.getProperties().add(i, new Property(HtmlConstants.CSS_BACKGROUND_COLOR, valueList, false));
            }
        } else {
            for (int j = 0; j < length; j++) {
                CSSValue item = valueList.item(j);
                String cssText = item.getCssText().toLowerCase();
                String color = Colors.fromStyle(cssText, null);
                if (color != null) {
                    cssStyleDeclaration.getProperties().add(i, new Property(HtmlConstants.CSS_BACKGROUND_COLOR, item, false));
                    break;
                }
            }
        }
    }

    private void splitBorder(CSSValueImpl valueList, int length, CSSStyleDeclarationImpl cssStyleDeclaration, int i) {
        if (length == 0) {
            String cssText = valueList.getCssText();
            if (StringUtils.isNotBlank(cssText)) {
                handleBorderValue(cssStyleDeclaration, i, valueList, cssText);
            }
        } else {
            for (int j = 0; j < length; j++) {
                CSSValue item = valueList.item(j);
                String value = item.getCssText();
                handleBorderValue(cssStyleDeclaration, i, item, value);
            }
        }
    }

    private void splitBorder(CSSValueImpl valueList, int length, CSSStyleDeclarationImpl cssStyleDeclaration, int i,
                             String styleProperty, String widthProperty, String colorProperty) {
        if (length == 0) {
            String cssText = valueList.getCssText();
            if (StringUtils.isNotBlank(cssText)) {
                handleBorderValue(cssStyleDeclaration, i, valueList, cssText, styleProperty, widthProperty, colorProperty);
            }
        } else {
            for (int j = 0; j < length; j++) {
                CSSValue item = valueList.item(j);
                String value = item.getCssText();
                handleBorderValue(cssStyleDeclaration, i, item, value, styleProperty, widthProperty, colorProperty);
            }
        }
    }

    private void handleBorderValue(CSSStyleDeclarationImpl cssStyleDeclaration, int i, CSSValue item, String value) {
        value = value.toLowerCase();
        if (HtmlConstants.BORDER_STYLES.contains(value)) {
            BoxProperty.BORDER_STYLE.setValues(cssStyleDeclaration, i, item);
        } else if (NamedBorderWidth.contains(value)) {
            BoxProperty.BORDER_WIDTH.setValues(cssStyleDeclaration, i, item);
        } else if (Character.isDigit(value.charAt(0))) {
            CSSLength width = CSSLength.of(value);
            if (width.isValid()) {
                BoxProperty.BORDER_WIDTH.setValues(cssStyleDeclaration, i, item);
            }
        } else {
            BoxProperty.BORDER_COLOR.setValues(cssStyleDeclaration, i, item);
        }
    }

    private void handleBorderValue(CSSStyleDeclarationImpl cssStyleDeclaration, int i, CSSValue item, String value,
                                   String styleProperty, String widthProperty, String colorProperty) {
        value = value.toLowerCase();
        if (HtmlConstants.BORDER_STYLES.contains(value)) {
            cssStyleDeclaration.getProperties().add(i, new Property(styleProperty, item, false));
        } else if (NamedBorderWidth.contains(value)) {
            cssStyleDeclaration.getProperties().add(i, new Property(widthProperty, item, false));
        } else if (Character.isDigit(value.charAt(0))) {
            CSSLength width = CSSLength.of(value);
            if (width.isValid()) {
                cssStyleDeclaration.getProperties().add(i, new Property(widthProperty, item, false));
            }
        } else {
            cssStyleDeclaration.getProperties().add(i, new Property(colorProperty, item, false));
        }
    }

    private void splitFont(CSSValueImpl valueList, int length, CSSStyleDeclarationImpl cssStyleDeclaration, int i) {
        if (length == 0) {
            return;
        }

        boolean styleHandled = false;
        boolean sizeHandled = false;
        for (int j = 0; j < length; j++) {
            CSSValue item = valueList.item(j);
            String value = item.getCssText();
            String lowerCase = value.toLowerCase();
            if (!styleHandled && HtmlConstants.FONT_STYLES.contains(lowerCase)) {
                cssStyleDeclaration.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_STYLE, item, false));
                styleHandled = true;
            } else if (HtmlConstants.FONT_VARIANTS.contains(lowerCase)) {
                cssStyleDeclaration.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_VARIANT_CAPS, item, false));
            } else if (HtmlConstants.FONT_WEIGHTS.contains(lowerCase) || NumberUtils.isParsable(value)) {
                cssStyleDeclaration.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_WEIGHT, item, false));
            } else if (HtmlConstants.SLASH.equals(value)) {
                // 字号与行高分隔符
                // https://www.w3.org/TR/CSS22/fonts.html#value-def-absolute-size
                // xx-small, x-small, small, medium, large, x-large, xx-large, xxx-large
                // 1,        ,        2,     3,      4,     5,       6,        7
                // FIXME font元素由于已废弃暂不支持
                // 长度/百分比
                CSSValue fontSize = valueList.item(j - 1);
                cssStyleDeclaration.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_SIZE, fontSize, false));
                sizeHandled = true;
                if (++j < length) {
                    // 数字/长度/百分比
                    CSSValue lineHeight = valueList.item(j);
                    cssStyleDeclaration.getProperties().add(i, new Property(HtmlConstants.CSS_LINE_HEIGHT, lineHeight, false));
                }
            } else if (HtmlConstants.COMMA.equals(value)) {
                // 多个字体之间的分隔符
                CSSValue firstFont = valueList.item(j - 1);
                if (!sizeHandled) {
                    CSSValue fontSize = valueList.item(j - 2);
                    cssStyleDeclaration.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_SIZE, fontSize, false));
                }
                if (HtmlConstants.isMajorFont(firstFont.getCssText())) {
                    cssStyleDeclaration.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_FAMILY, firstFont, false));
                } else {
                    for (j++; j < length; j++) {
                        CSSValue fontFamily = valueList.item(j);
                        if (HtmlConstants.isMajorFont(fontFamily.getCssText())) {
                            cssStyleDeclaration.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_FAMILY, fontFamily, false));
                            break;
                        }
                    }
                }
                break;
            } else if (j == length - 1) {
                // font-family在font中一定是最后出现
                cssStyleDeclaration.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_FAMILY, item, false));
            }
        }
    }

    private void splitBox(CSSValueImpl valueList, int length, CSSStyleDeclarationImpl cssStyleDeclaration, int i,
                          BoxProperty boxProperty) {
        switch (length) {
            // 当仅一个值时实际返回长度为0
            case 0:
            case 1:
                if (StringUtils.isNotBlank(valueList.getCssText())) {
                    boxProperty.setValues(cssStyleDeclaration, i, valueList);
                }
                break;
            case 2:
                boxProperty.setValues(cssStyleDeclaration, i, valueList.item(0), valueList.item(1));
                break;
            case 3:
                boxProperty.setValues(cssStyleDeclaration, i, valueList.item(0), valueList.item(1), valueList.item(2));
                break;
            case 4:
                boxProperty.setValues(cssStyleDeclaration, i, valueList.item(0), valueList.item(1), valueList.item(2), valueList.item(3));
                break;
        }
    }

}

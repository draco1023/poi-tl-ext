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
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.BodyType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
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
import org.ddr.poi.html.tag.SvgRenderer;
import org.ddr.poi.html.tag.TableCellRenderer;
import org.ddr.poi.html.tag.TableRenderer;
import org.ddr.poi.html.tag.UnderlineRenderer;
import org.ddr.poi.html.tag.WalkThroughRenderer;
import org.ddr.poi.html.util.CSSLength;
import org.ddr.poi.html.util.CSSStyleUtils;
import org.ddr.poi.html.util.JsoupUtils;
import org.ddr.poi.html.util.RenderUtils;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
    private final HtmlRenderConfig config;

    public HtmlRenderPolicy() {
        this(new HtmlRenderConfig());
    }

    @Deprecated
    public HtmlRenderPolicy(String globalFont, CSSLength globalFontSize) {
        this(new HtmlRenderConfig());
        config.setGlobalFont(globalFont);
        config.setGlobalFontSize(globalFontSize);
    }

    public HtmlRenderPolicy(HtmlRenderConfig config) {
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
                new SvgRenderer(),
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
        this.config = config;
        // custom tag renderer will overwrite the built-in renderer
        if (config.getCustomRenderers() != null) {
            for (ElementRenderer customRenderer : config.getCustomRenderers()) {
                for (String tag : customRenderer.supportedTags()) {
                    elRenderers.put(tag, customRenderer);
                }
            }
        }
    }

    public HtmlRenderConfig getConfig() {
        return config;
    }

    @Override
    protected boolean validate(String data) {
        return StringUtils.isNotEmpty(data);
    }

    @Override
    public void doRender(RenderContext<String> context) throws Exception {
        String html = FORMATTED_PATTERN.matcher(context.getData()).replaceAll(FORMATTED_REPLACEMENT);
        Document document = JsoupUtils.parseBodyFragment(html);
        document.outputSettings().prettyPrint(false).indentAmount(0);

        HtmlRenderContext htmlRenderContext = new HtmlRenderContext(context);
        htmlRenderContext.setGlobalFont(config.getGlobalFont());
        if (config.getGlobalFontSizeInHalfPoints() > 0) {
            htmlRenderContext.setGlobalFontSize(BigInteger.valueOf(config.getGlobalFontSizeInHalfPoints()));
        }
        htmlRenderContext.getNumberingContext().setIndent(config.getNumberingIndent());
        htmlRenderContext.getNumberingContext().setSpacing(config.getNumberingSpacing());

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
            context.renderText(((TextNode) node).getWholeText());
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
        boolean blocked = false;

        if (element.isBlock() && (elementRenderer == null || elementRenderer.renderAsBlock())) {
            if (element.childNodeSize() == 0 && !HtmlConstants.KEEP_EMPTY_TAGS.contains(element.normalName())) {
                return;
            }
            if (!context.isBlocked()) {
                // 复制段落中占位符之前的部分内容
                moveContentToNewPrevParagraph(context);
            }
            context.incrementBlockLevel();
            blocked = true;

            IBody container = context.getContainer();
            boolean isTableTag = HtmlConstants.TAG_TABLE.equals(element.normalName());

            XmlCursor xmlCursor = getElementCursor(context, container, isTableTag);
            if (xmlCursor == null) {
                context.decrementBlockLevel();
                return;
            }

            if (isTableTag) {
                XWPFTable xwpfTable = container.insertNewTbl(xmlCursor);
                xmlCursor.dispose();
                // 新增时会自动创建一行一列，会影响自定义的表格渲染逻辑，故删除
                xwpfTable.removeRow(0);
                context.replaceClosestBody(xwpfTable);

                if (container.getPartType() == BodyType.TABLECELL && config.isShowDefaultTableBorderInTableCell()) {
                    CTTbl ctTbl = xwpfTable.getCTTbl();
                    CTTblPr tblPr = RenderUtils.getTblPr(ctTbl);
                    CTTblBorders tblBorders = RenderUtils.getTblBorders(tblPr);
                    tblBorders.addNewTop().setVal(STBorder.SINGLE);
                    tblBorders.addNewLeft().setVal(STBorder.SINGLE);
                    tblBorders.addNewBottom().setVal(STBorder.SINGLE);
                    tblBorders.addNewRight().setVal(STBorder.SINGLE);
                    tblBorders.addNewInsideH().setVal(STBorder.SINGLE);
                    tblBorders.addNewInsideV().setVal(STBorder.SINGLE);
                }

                RenderUtils.tableStyle(context, xwpfTable, cssStyleDeclaration);
            } else if (shouldNewParagraph(element)) {
                XWPFParagraph xwpfParagraph = context.newParagraph(container, xmlCursor);
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
                renderElementEnd(element, context, elementRenderer, blocked);
                return;
            }
        }

        for (Node child : element.childNodes()) {
            renderNode(child, context);
        }

        renderElementEnd(element, context, elementRenderer, blocked);
    }

    private boolean shouldNewParagraph(Element element) {
        // li的第一个子节点如果为块状元素，避免生成新的段落
        if (element.hasParent() && HtmlConstants.TAG_LI.equals(element.parent().normalName())
                && element.parentNode().childNode(0) == element) {
            return false;
        }
        return true;
    }

    private XmlCursor getElementCursor(HtmlRenderContext context, IBody container, boolean isTableTag) {
        XmlCursor xmlCursor;
        if (context.containerChanged()) {
            IBodyElement closestBody = context.getClosestBody();
            switch (closestBody.getElementType()) {
                case PARAGRAPH:
                    xmlCursor = ((XWPFParagraph) closestBody).getCTP().newCursor();
                    xmlCursor.toEndToken();
                    xmlCursor.toNextToken();
                    break;
                case TABLE:
                    xmlCursor = ((XWPFTable) closestBody).getCTTbl().newCursor();
                    xmlCursor.toEndToken();
                    xmlCursor.toNextToken();
                    if (isTableTag) {
                        // 插入一个段落，防止表格粘连在一起
                        context.newParagraph(container, xmlCursor);
                        xmlCursor.toNextToken();
                    }
                    break;
                default:
                    return null;
            }
        } else {
            xmlCursor = context.getRun().getCTR().newCursor();
            xmlCursor.toParent();
            xmlCursor.push();
            // 如果是表格，检查当前word容器的前一个兄弟元素是否为表格，是则插入一个段落，防止表格粘连在一起
            if (isTableTag && xmlCursor.toPrevSibling()) {
                if (xmlCursor.getObject() instanceof CTTbl) {
                    xmlCursor.toNextSibling();
                    context.newParagraph(container, xmlCursor);
                }
            }
            xmlCursor.pop();
        }
        return xmlCursor;
    }

    private void renderElementEnd(Element element, HtmlRenderContext context, ElementRenderer elementRenderer, boolean blocked) {
        if (elementRenderer != null) {
            elementRenderer.renderEnd(element, context);
        }
        context.popInlineStyle();
        if (blocked) {
            context.decrementBlockLevel();
        }
    }

    private void moveContentToNewPrevParagraph(HtmlRenderContext context) {
        CTR ctr = context.getRun().getCTR();
        XmlCursor rCursor = ctr.newCursor();
        boolean hasPrevSibling = false;
        while (rCursor.toPrevSibling()) {
            if (!(rCursor.getObject() instanceof CTPPr)) {
                hasPrevSibling = true;
                break;
            }
        }
        if (!hasPrevSibling) {
            rCursor.dispose();
            return;
        }
        rCursor.toParent();
        rCursor.push();
        XWPFParagraph newParagraph = context.getContainer().insertNewParagraph(rCursor);
        XmlCursor pCursor = newParagraph.getCTP().newCursor();
        pCursor.toEndToken();
        rCursor.pop();
        rCursor.toFirstChild();
        while (!ctr.equals(rCursor.getObject())) {
            XmlObject obj = rCursor.getObject();
            if (obj instanceof CTPPr) {
                rCursor.copyXml(pCursor);
                rCursor.toNextSibling();
            } else if (obj instanceof CTBookmark) {
                rCursor.toNextSibling();
            } else {
                // moveXml附带了toNextSibling的效果
                rCursor.moveXml(pCursor);
            }
        }
        rCursor.dispose();
        pCursor.dispose();
    }

    @Override
    protected void afterRender(RenderContext<String> context) {
        boolean hasSibling = hasSibling(context.getRun());
        clearPlaceholder(context, !hasSibling);

        IBody container = context.getContainer();
        if (container.getPartType() == BodyType.TABLECELL) {
            // 单元格的最后一个元素应为p，否则可能无法正常打开文件
            List<IBodyElement> bodyElements = container.getBodyElements();
            if (bodyElements.isEmpty() || bodyElements.get(bodyElements.size() - 1).getElementType() != BodyElementType.PARAGRAPH) {
                ((XWPFTableCell) container).addParagraph();
            }
        }
    }

    private boolean hasSibling(XWPFRun run) {
        boolean hasSibling = false;
        CTR ctr = run.getCTR();
        XmlCursor xmlCursor = ctr.newCursor();
        xmlCursor.push();
        while (xmlCursor.toNextSibling()) {
            if (isValidSibling(xmlCursor.getObject())) {
                hasSibling = true;
                break;
            }
        }
        if (!hasSibling) {
            xmlCursor.pop();
            while (xmlCursor.toPrevSibling()) {
                if (isValidSibling(xmlCursor.getObject())) {
                    hasSibling = true;
                    break;
                }
            }
        }
        xmlCursor.dispose();
        return hasSibling;
    }

    private boolean isValidSibling(XmlObject object) {
        return !(object instanceof CTPPr) && !(object instanceof CTBookmark);
    }

    private CSSStyleDeclarationImpl getCssStyleDeclaration(Element element) {
        String style = element.attr(HtmlConstants.ATTR_STYLE);
        CSSStyleDeclarationImpl cssStyleDeclaration = CSSStyleUtils.parse(style);
        CSSStyleUtils.split(cssStyleDeclaration);
        return cssStyleDeclaration;
    }

}

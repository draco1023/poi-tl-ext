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
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.BodyType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.ddr.poi.html.tag.ARenderer;
import org.ddr.poi.html.tag.BigRenderer;
import org.ddr.poi.html.tag.BoldRenderer;
import org.ddr.poi.html.tag.BreakRenderer;
import org.ddr.poi.html.tag.DeleteRenderer;
import org.ddr.poi.html.tag.FigureCaptionRenderer;
import org.ddr.poi.html.tag.FigureRenderer;
import org.ddr.poi.html.tag.HeaderBreakRenderer;
import org.ddr.poi.html.tag.HeaderRenderer;
import org.ddr.poi.html.tag.ImageRenderer;
import org.ddr.poi.html.tag.ItalicRenderer;
import org.ddr.poi.html.tag.ListItemRenderer;
import org.ddr.poi.html.tag.ListRenderer;
import org.ddr.poi.html.tag.MarkRenderer;
import org.ddr.poi.html.tag.MathRenderer;
import org.ddr.poi.html.tag.OmittedRenderer;
import org.ddr.poi.html.tag.RubyRenderer;
import org.ddr.poi.html.tag.SmallRenderer;
import org.ddr.poi.html.tag.SubscriptRenderer;
import org.ddr.poi.html.tag.SuperscriptRenderer;
import org.ddr.poi.html.tag.SvgRenderer;
import org.ddr.poi.html.tag.TableCellRenderer;
import org.ddr.poi.html.tag.TableRenderer;
import org.ddr.poi.html.tag.UnderlineRenderer;
import org.ddr.poi.html.tag.WalkThroughRenderer;
import org.ddr.poi.html.util.CSSLength;
import org.ddr.poi.html.util.JsoupUtils;
import org.jsoup.nodes.Document;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

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
                new FigureRenderer(),
                new FigureCaptionRenderer(),
                new HeaderBreakRenderer(),
                new HeaderRenderer(),
                new ImageRenderer(),
                new ItalicRenderer(),
                new ListItemRenderer(),
                new ListRenderer(),
                new MarkRenderer(),
                new MathRenderer(),
                new OmittedRenderer(),
                new RubyRenderer(),
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
        Document document = JsoupUtils.parse(html);
        document.outputSettings().prettyPrint(false).indentAmount(0);

        HtmlRenderContext htmlRenderContext = new HtmlRenderContext(context, elRenderers::get);
        htmlRenderContext.setGlobalFont(config.getGlobalFont());
        if (config.getGlobalFontSizeInHalfPoints() > 0) {
            htmlRenderContext.setGlobalFontSize(BigInteger.valueOf(config.getGlobalFontSizeInHalfPoints()));
        }
        htmlRenderContext.getNumberingContext().setIndent(config.getNumberingIndent());
        htmlRenderContext.getNumberingContext().setSpacing(config.getNumberingSpacing());
        htmlRenderContext.setShowDefaultTableBorderInTableCell(config.isShowDefaultTableBorderInTableCell());

        htmlRenderContext.renderDocument(document);
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

}

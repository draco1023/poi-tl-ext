/*
 * Copyright 2016 - 2022 Draco, https://github.com/draco1023
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

package org.ddr.poi.html.tag;

import org.apache.commons.lang3.StringUtils;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.ddr.poi.html.util.RenderUtils;
import org.jsoup.internal.StringUtil;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;

/**
 * ruby标签渲染器
 *
 * @author Draco
 * @since 2022-06-12 20:47
 */
public class RubyRenderer implements ElementRenderer {
    private static final String[] TAGS = {
            HtmlConstants.TAG_RUBY,
    };

    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        StringBuilder sb = StringUtil.borrowBuilder();
        for (Node childNode : element.childNodes()) {
            if (childNode instanceof Element) {
                String tagName = ((Element) childNode).normalName();
                if (HtmlConstants.TAG_RT.equals(tagName)) {
                    String rt = ((Element) childNode).text();

                    CTR ctr = context.newRun();
                    CTRPr rPr = RenderUtils.getRPr(ctr);
                    rPr.addNewLang().setVal("en-US");
                    ctr.addNewFldChar().setFldCharType(STFldCharType.BEGIN);

                    ctr = context.newRun();
                    rPr = RenderUtils.getRPr(ctr);
                    rPr.addNewLang().setVal("en-US");
                    CTText ctText = ctr.addNewInstrText();
                    ctText.setSpace(SpaceAttribute.Space.PRESERVE);

                    int fontSize = context.getGlobalFontSize() == null
                            ? context.getInheritedFontSizeInHalfPoints() : context.getGlobalFontSize().intValue();
                    fontSize = (fontSize + 1) / 2;
                    ctText.setStringValue("EQ \\* jc0 \\* hps" + fontSize + " \\o \\ad(\\s \\up 9(" + rt + "),"
                            + StringUtil.releaseBuilder(sb).trim() + ")");
                    sb = StringUtil.borrowBuilder();

                    ctr = context.newRun();
                    rPr = RenderUtils.getRPr(ctr);
                    rPr.addNewLang().setVal("en-US");
                    ctr.addNewFldChar().setFldCharType(STFldCharType.END);
                } else if (HtmlConstants.TAG_RP.equals(tagName)) {
                    continue;
                } else {
                    StringUtil.appendNormalisedWhitespace(sb, ((Element) childNode).wholeText(), false);
                }
            } else if (childNode instanceof TextNode) {
                StringUtil.appendNormalisedWhitespace(sb, ((TextNode) childNode).getWholeText(), false);
            }
        }
        String remainText = StringUtil.releaseBuilder(sb);
        if (StringUtils.isNotBlank(remainText)) {
            context.renderText(remainText);
        }
        return false;
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

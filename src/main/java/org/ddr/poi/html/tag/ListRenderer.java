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

package org.ddr.poi.html.tag;

import org.apache.commons.lang3.StringUtils;
import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.ddr.poi.html.util.CSSLength;
import org.ddr.poi.html.util.ListStyle;
import org.ddr.poi.html.util.ListStyleType;
import org.ddr.poi.html.util.RenderUtils;
import org.jsoup.nodes.Element;

/**
 * 列表渲染器
 *
 * @author Draco
 * @since 2021-02-18
 */
public class ListRenderer implements ElementRenderer {
    private static final String[] TAGS = {HtmlConstants.TAG_UL, HtmlConstants.TAG_OL};

    /**
     * 开始渲染
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @return 是否继续渲染子元素
     */
    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        String listStylePosition = context.currentElementStyle().getPropertyValue(HtmlConstants.CSS_LIST_STYLE_POSITION);
        boolean hanging = !HtmlConstants.INSIDE.equals(listStylePosition);
        CSSLength marginLeft = CSSLength.of(context.currentElementStyle().getMarginLeft().toLowerCase());
        int left = marginLeft.isValid() && !marginLeft.isPercent()
                ? RenderUtils.emuToTwips(context.lengthToEMU(marginLeft)) : 0;
        CSSLength marginRight = CSSLength.of(context.currentElementStyle().getMarginRight().toLowerCase());
        int right = marginRight.isValid() && !marginRight.isPercent()
                ? RenderUtils.emuToTwips(context.lengthToEMU(marginRight)) : 0;
        ListStyle listStyle = new ListStyle(determineNumberFormat(context, element), hanging, left, right);
        context.getNumberingContext().startLevel(listStyle);
        return true;
    }

    private ListStyleType determineNumberFormat(HtmlRenderContext context, Element element) {
        String listStyleType = context.currentElementStyle()
                .getPropertyValue(HtmlConstants.CSS_LIST_STYLE_TYPE).toLowerCase();
        ListStyleType format;
        switch (element.tag().normalName()) {
            case HtmlConstants.TAG_OL:
                if (StringUtils.isNotBlank(listStyleType)) {
                    format = ListStyleType.Ordered.of(listStyleType);
                } else {
                    // 支持ol的type属性
                    String type = element.attr(HtmlConstants.ATTR_TYPE);
                    format = ListStyleType.Ordered.of(type);
                }
                break;
            case HtmlConstants.TAG_UL:
                format = ListStyleType.Unordered.of(listStyleType);
                break;
            default:
                format = ListStyleType.Unordered.NONE;
        }
        return format;
    }

    /**
     * 元素渲染结束需要执行的逻辑
     *
     * @param element HTML元素
     * @param context 渲染上下文
     */
    @Override
    public void renderEnd(Element element, HtmlRenderContext context) {
        context.getNumberingContext().endLevel();
    }

    @Override
    public String[] supportedTags() {
        return TAGS;
    }

    @Override
    public boolean renderAsBlock() {
        // 列表标签本身不需要作为块状元素渲染，因为每一个列表项都是一个块状元素
        return false;
    }
}

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

import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.jsoup.nodes.Element;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;

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
        context.getNumberingContext().startLevel(determineNumberFormat(element));
        return true;
    }

    private STNumberFormat.Enum determineNumberFormat(Element element) {
        // TODO 支持ol的type属性以及css的list-style-type
        STNumberFormat.Enum format;
        switch (element.tag().normalName()) {
            case HtmlConstants.TAG_OL:
                format = STNumberFormat.DECIMAL;
                break;
            case HtmlConstants.TAG_UL:
                format = STNumberFormat.BULLET;
                break;
            default:
                format = STNumberFormat.NONE;
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

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

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.jsoup.nodes.Element;

/**
 * 列表项渲染器
 *
 * @author Draco
 * @since 2021-02-19
 */
public class ListItemRenderer implements ElementRenderer {
    private static final String[] TAGS = {HtmlConstants.TAG_LI};

    /**
     * 开始渲染
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @return 是否继续渲染子元素
     */
    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        XWPFParagraph paragraph = context.getClosestParagraph();
        context.markDedupe(paragraph);
        context.getNumberingContext().add(paragraph);
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

    @Override
    public String[] supportedTags() {
        return TAGS;
    }

    @Override
    public boolean renderAsBlock() {
        return true;
    }
}

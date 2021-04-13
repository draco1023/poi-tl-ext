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
import org.ddr.poi.html.util.RenderUtils;
import org.jsoup.nodes.Element;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPBdr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;

import java.math.BigInteger;

/**
 * hr标签渲染器
 *
 * @author Draco
 * @since 2021-02-18
 */
public class HeaderBreakRenderer implements ElementRenderer {
    private static final String[] TAGS = {HtmlConstants.TAG_HR};
    /**
     * 线粗细，相当于3px
     */
    private static final BigInteger SIZE = BigInteger.valueOf(12);
    /**
     * 间距
     */
    private static final BigInteger SPACE = BigInteger.ONE;

    /**
     * 开始渲染
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @return 是否继续渲染子元素
     */
    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        CTP ctp = context.getClosestParagraph().getCTP();
        CTPBdr pBdr = RenderUtils.getPBdr(RenderUtils.getPPr(ctp));
        CTBorder ctBorder = pBdr.addNewBottom();
        ctBorder.setVal(STBorder.SINGLE);
        ctBorder.setSz(SIZE);
        ctBorder.setSpace(SPACE);
//        ctBorder.setColor("auto");
        return false;
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

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

import org.jsoup.nodes.Element;


/**
 * HTML元素渲染器
 *
 * @author Draco
 * @since 2021-02-08
 */
public interface ElementRenderer {
    /**
     * 开始渲染
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @return 是否继续渲染子元素
     */
    boolean renderStart(Element element, HtmlRenderContext context);

    /**
     * 元素渲染结束需要执行的逻辑
     *
     * @param element HTML元素
     * @param context 渲染上下文
     */
    default void renderEnd(Element element, HtmlRenderContext context) {
    }

    /**
     * @return 支持的HTML标签
     */
    String[] supportedTags();

    /**
     * @return 是否为块状渲染，如果为true在Word中会另起一个Paragraph
     */
    boolean renderAsBlock();
}

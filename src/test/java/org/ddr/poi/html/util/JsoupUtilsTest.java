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

package org.ddr.poi.html.util;

import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

/**
* @author Korben
* @Created on: 2026-08-24 10:40
* @since:
* @description: JsoupUtils解析测试，验证SVG大小写保留等HTML解析行为不随jsoup版本升级而回归
*/
class JsoupUtilsTest {

    @Test
    void parsePreservesSvgCase() {
        String html = "<div><svg viewBox=\"0 0 100 100\" preserveAspectRatio=\"xMidYMid meet\">"
                + "<linearGradient id=\"grad\"><stop offset=\"0\" stop-color=\"red\"/></linearGradient>"
                + "<foreignObject><p>text</p></foreignObject></svg></div>";
        Document document = JsoupUtils.parse(html);

        Element svg = document.selectFirst("svg");
        assertNotNull(svg);
        // 属性名大小写保留
        assertEquals("0 0 100 100", svg.attr("viewBox"));
        assertEquals("xMidYMid meet", svg.attr("preserveAspectRatio"));
        // 子标签名大小写保留
        Element gradient = svg.child(0);
        assertEquals("linearGradient", gradient.tagName());
        Element foreignObject = svg.child(1);
        assertEquals("foreignObject", foreignObject.tagName());
        // svg内的HTML内容（foreignObject）正常解析
        assertEquals("text", foreignObject.text());
    }

    @Test
    void parseNormalizesHtmlCase() {
        Document document = JsoupUtils.parse("<DIV CLASS=\"Box\">Text</DIV>");

        Element div = document.selectFirst("div");
        assertNotNull(div);
        assertEquals("div", div.tagName());
        assertEquals("Box", div.attr("class"));
        assertEquals("Text", div.text());
    }

}

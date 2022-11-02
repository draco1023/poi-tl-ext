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

import org.ddr.poi.html.HtmlConstants;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.parser.CustomHtmlTreeBuilder;
import org.jsoup.parser.Parser;
import org.jsoup.select.Elements;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;
import java.util.function.Predicate;

/**
 * JSoup工具类
 *
 * @author Draco
 * @since 2021-03-03
 */
public class JsoupUtils {
    /**
     * 选取符合条件的子元素到目标集合中
     *
     * @param collection 目标集合
     * @param parent 父元素
     * @param predicate 条件
     */
    public static void selectChildren(Elements collection, Element parent, Predicate<Element> predicate) {
        for (Node node : parent.childNodes()) {
            if (node instanceof Element) {
                Element child = ((Element) node);
                if (predicate.test(child)) {
                    collection.add(child);
                }
            }
        }
    }

    /**
     * 选取指定标签的子元素
     *
     * @param parent 父元素
     * @param tag 标签名称，小写
     * @return 子元素集合
     */
    public static Elements children(Element parent, String tag) {
        Elements elements = new Elements();
        selectChildren(elements, parent, c -> c.normalName().equals(tag));
        return elements;
    }

    /**
     * 选取指定标签的子元素
     *
     * @param parent 父元素
     * @param tags 多种标签名称，小写
     * @return 子元素集合
     */
    public static Elements children(Element parent, String... tags) {
        Elements elements = new Elements();
        Set<String> targets = new HashSet<>(Arrays.asList(tags));
        selectChildren(elements, parent, c -> targets.contains(c.normalName()));
        return elements;
    }

    /**
     * 选取第一个指定标签的子元素
     *
     * @param parent 父元素
     * @param tag 标签名称，小写
     * @return 子元素
     */
    public static Element firstChild(Element parent, String tag) {
        for (Node node : parent.childNodes()) {
            if (node instanceof Element) {
                Element child = ((Element) node);
                if (child.normalName().equals(tag)) {
                    return child;
                }
            }
        }
        return null;
    }

    /**
     * 选取表格的所有行元素
     *
     * @param parent 表格元素
     * @return 行元素集合
     */
    public static Elements childRows(Element parent) {
        Elements elements = new Elements();
        for (Node node : parent.childNodes()) {
            if (node instanceof Element) {
                Element child = ((Element) node);
                if (HtmlConstants.TAG_TR.equals(child.normalName())) {
                    // 直接位于table标签下
                    elements.add(child);
                } else {
                    // 可能位于thead/tbody/tfoot标签下，选取直接子元素避免受嵌套表格影响
                    selectChildren(elements, child, c -> HtmlConstants.TAG_TR.equals(c.normalName()));
                }
            }
        }
        return elements;
    }

    /**
     * @see org.jsoup.Jsoup#parseBodyFragment(String)
     * @see org.jsoup.parser.Parser#parseBodyFragment(String, String)
     */
    public static Document parse(String html) {
        CustomHtmlTreeBuilder treeBuilder = new CustomHtmlTreeBuilder();
        return Jsoup.parse(html, new Parser(treeBuilder));
    }

}

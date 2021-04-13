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

import com.steadystate.css.dom.CSSStyleDeclarationImpl;
import com.steadystate.css.dom.Property;
import org.ddr.poi.html.HtmlConstants;
import org.w3c.dom.css.CSSValue;

/**
 * CSS中支持四个边的属性
 *
 * @author Draco
 * @since 2021-03-29
 */
public enum BoxProperty {
    MARGIN(HtmlConstants.CSS_MARGIN_TOP, HtmlConstants.CSS_MARGIN_RIGHT,
            HtmlConstants.CSS_MARGIN_BOTTOM, HtmlConstants.CSS_MARGIN_LEFT),
    PADDING(HtmlConstants.CSS_PADDING_TOP, HtmlConstants.CSS_PADDING_RIGHT,
            HtmlConstants.CSS_PADDING_BOTTOM, HtmlConstants.CSS_PADDING_LEFT),
    BORDER_STYLE(HtmlConstants.CSS_BORDER_TOP_STYLE, HtmlConstants.CSS_BORDER_RIGHT_STYLE,
            HtmlConstants.CSS_BORDER_BOTTOM_STYLE, HtmlConstants.CSS_BORDER_LEFT_STYLE),
    BORDER_WIDTH(HtmlConstants.CSS_BORDER_TOP_WIDTH, HtmlConstants.CSS_BORDER_RIGHT_WIDTH,
            HtmlConstants.CSS_BORDER_BOTTOM_WIDTH, HtmlConstants.CSS_BORDER_LEFT_WIDTH),
    BORDER_COLOR(HtmlConstants.CSS_BORDER_TOP_COLOR, HtmlConstants.CSS_BORDER_RIGHT_COLOR,
            HtmlConstants.CSS_BORDER_BOTTOM_COLOR, HtmlConstants.CSS_BORDER_LEFT_COLOR);

    private final String top;
    private final String right;
    private final String bottom;
    private final String left;

    BoxProperty(String top, String right, String bottom, String left) {
        this.top = top;
        this.right = right;
        this.bottom = bottom;
        this.left = left;
    }

    public void setValues(CSSStyleDeclarationImpl cssStyleDeclaration, int i,
                          CSSValue topValue, CSSValue rightValue, CSSValue bottomValue, CSSValue leftValue) {
        cssStyleDeclaration.getProperties().add(i, new Property(top, topValue, false));
        cssStyleDeclaration.getProperties().add(i, new Property(right, rightValue, false));
        cssStyleDeclaration.getProperties().add(i, new Property(bottom, bottomValue, false));
        cssStyleDeclaration.getProperties().add(i, new Property(left, leftValue, false));
    }

    public void setValues(CSSStyleDeclarationImpl cssStyleDeclaration, int i, CSSValue value) {
        setValues(cssStyleDeclaration, i, value, value, value, value);
    }

    public void setValues(CSSStyleDeclarationImpl cssStyleDeclaration, int i, CSSValue topBottom, CSSValue rightLeft) {
        setValues(cssStyleDeclaration, i, topBottom, rightLeft, topBottom, rightLeft);
    }

    public void setValues(CSSStyleDeclarationImpl cssStyleDeclaration, int i, CSSValue top, CSSValue rightLeft, CSSValue bottom) {
        setValues(cssStyleDeclaration, i, top, rightLeft, bottom, rightLeft);
    }
}

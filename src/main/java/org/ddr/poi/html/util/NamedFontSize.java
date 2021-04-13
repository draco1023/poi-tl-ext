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

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 已命名的字号
 *
 * @author Draco
 * @since 2021-02-25
 */
public enum NamedFontSize {
    XX_SMALL(HtmlConstants.XX_SMALL, 6d),
    X_SMALL(HtmlConstants.X_SMALL, 7.5d),
    SMALL(HtmlConstants.SMALL, 10d),
    MEDIUM(HtmlConstants.MEDIUM, 12d),
    LARGE(HtmlConstants.LARGE, 13.5d),
    X_LARGE(HtmlConstants.X_LARGE, 18d),
    XX_LARGE(HtmlConstants.XX_LARGE, 24d),
    XXX_LARGE(HtmlConstants.XXX_LARGE, 36d);

    private final String name;
    private final CSSLength size;
    private static final Map<String, NamedFontSize> NAMED_MAP = Arrays.stream(NamedFontSize.values()).collect(
            Collectors.toMap(NamedFontSize::getName, Function.identity())
    );

    NamedFontSize(String name, double pt) {
        this.name = name;
        this.size = new CSSLength(pt, CSSLengthUnit.PT);
    }

    public String getName() {
        return name;
    }

    public CSSLength getSize() {
        return size;
    }

    public static NamedFontSize of(String name) {
        return NAMED_MAP.get(name);
    }
}

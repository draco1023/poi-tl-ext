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
 * 已命名的边框宽度，宽度值与chrome一致
 *
 * @author Draco
 * @since 2021-03-29
 */
public enum NamedBorderWidth {
    THIN(HtmlConstants.THIN, 1),
    MEDIUM(HtmlConstants.MEDIUM, 3),
    THICK(HtmlConstants.THICK, 5);

    private final String name;
    private final CSSLength width;
    private static final Map<String, NamedBorderWidth> NAMED_MAP = Arrays.stream(NamedBorderWidth.values()).collect(
            Collectors.toMap(NamedBorderWidth::getName, Function.identity())
    );

    NamedBorderWidth(String name, int px) {
        this.name = name;
        this.width = new CSSLength(px, CSSLengthUnit.PX);
    }

    public String getName() {
        return name;
    }

    public CSSLength getWidth() {
        return width;
    }

    public static NamedBorderWidth of(String name) {
        return NAMED_MAP.get(name);
    }

    public static boolean contains(String name) {
        return NAMED_MAP.containsKey(name);
    }
}

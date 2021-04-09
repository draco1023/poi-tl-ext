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

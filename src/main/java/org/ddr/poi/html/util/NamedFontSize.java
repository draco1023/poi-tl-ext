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

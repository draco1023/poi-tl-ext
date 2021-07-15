package org.ddr.poi.html.util;

import org.ddr.poi.html.HtmlConstants;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * https://developer.mozilla.org/zh-CN/docs/Web/CSS/white-space
 *
 * @author Draco
 * @since 2021-07-15
 */
public enum WhiteSpaceRule {
    NORMAL(HtmlConstants.NORMAL, false, false, false),
    NO_WRAP(HtmlConstants.NO_WRAP, false, false, false),
    PRE(HtmlConstants.PRE, true, true, true),
    PRE_WRAP(HtmlConstants.PRE_WRAP, true, true, true),
    PRE_LINE(HtmlConstants.PRE_LINE, true, false, false),
    BREAK_SPACES(HtmlConstants.BREAK_SPACES, true, true, true);

    private final String value;
    private final boolean keepLineBreak;
    private final boolean keepSpaceAndTab;
    private final boolean keepTrailingSpace;

    WhiteSpaceRule(String value, boolean keepLineBreak, boolean keepSpaceAndTab, boolean keepTrailingSpace) {
        this.value = value;
        this.keepLineBreak = keepLineBreak;
        this.keepSpaceAndTab = keepSpaceAndTab;
        this.keepTrailingSpace = keepTrailingSpace;
    }

    public String getValue() {
        return value;
    }

    public boolean isKeepLineBreak() {
        return keepLineBreak;
    }

    public boolean isKeepSpaceAndTab() {
        return keepSpaceAndTab;
    }

    public boolean isKeepTrailingSpace() {
        return keepTrailingSpace;
    }

    private static Map<String, WhiteSpaceRule> rules = Arrays.stream(values()).collect(
            Collectors.toMap(WhiteSpaceRule::getValue, Function.identity())
    );

    public static WhiteSpaceRule of(String value) {
        return rules.getOrDefault(value, WhiteSpaceRule.NORMAL);
    }
}

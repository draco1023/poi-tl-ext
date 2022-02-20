package org.ddr.poi.html.util;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 列表项样式
 *
 * @author Draco
 * @since 2022-02-08
 */
public interface ListStyleType {
    String SYMBOL_DISC = "\uF06C"; // •⚫
    String SYMBOL_CIRCLE = "\uF0A1"; // ◦⚪
    String SYMBOL_DISCLOSURE_CLOSED = "\uF075"; // ▸
    String SYMBOL_DISCLOSURE_OPEN = "\uF071"; // ▾
    String SYMBOL_SQUARE = "\uF06E"; // ▪

    String FONT_WINGDINGS = "Wingdings";
    String FONT_WINGDINGS_3 = "Wingdings 3";

    String getName();

    STNumberFormat.Enum getFormat();

    /**
     * string: Custom symbol empty string: System counting style null: No symbol
     */
    String getText();

    String getFont();

    enum Unordered implements ListStyleType {
        DISC("disc", STNumberFormat.BULLET, SYMBOL_DISC, FONT_WINGDINGS),
        CIRCLE("circle", STNumberFormat.BULLET, SYMBOL_CIRCLE, FONT_WINGDINGS),
        DECIMAL("decimal", STNumberFormat.DECIMAL, "", null),
        DISCLOSURE_CLOSED("disclosure-closed", STNumberFormat.BULLET, SYMBOL_DISCLOSURE_CLOSED, FONT_WINGDINGS_3),
        DISCLOSURE_OPEN("disclosure-open", STNumberFormat.BULLET, SYMBOL_DISCLOSURE_OPEN, FONT_WINGDINGS_3),
        SQUARE("square", STNumberFormat.BULLET, SYMBOL_SQUARE, FONT_WINGDINGS),
        NONE("none", STNumberFormat.NONE, null, null);

        private static final Map<String, ListStyleType> TYPE_MAP = Arrays.stream(values())
                .collect(Collectors.toMap(Unordered::getName, Function.identity()));

        private final String name;
        private final STNumberFormat.Enum format;
        private final String text;
        private final String font;

        Unordered(String name, STNumberFormat.Enum format, String text, String font) {
            this.name = name;
            this.format = format;
            this.text = text;
            this.font = font;
        }

        @Override
        public String getName() {
            return name;
        }

        @Override
        public STNumberFormat.Enum getFormat() {
            return format;
        }

        @Override
        public String getText() {
            return text;
        }

        @Override
        public String getFont() {
            return font;
        }

        public static ListStyleType of(String type) {
            return TYPE_MAP.getOrDefault(type, DISC);
        }
    }

    /**
     * https://www.w3.org/TR/css-counter-styles-3/#predefined-counters
     */
    enum Ordered implements ListStyleType {
        /* Numeric */
        DECIMAL("decimal", STNumberFormat.DECIMAL, "", null),
        DECIMAL_LEADING_ZERO("decimal-leading-zero", STNumberFormat.DECIMAL_ZERO, "", null),
        CJK_DECIMAL("cjk-decimal", STNumberFormat.TAIWANESE_DIGITAL, "", null),
        HEBREW("hebrew", STNumberFormat.HEBREW_1, "", null),
        LOWER_ROMAN("lower-roman", STNumberFormat.LOWER_ROMAN, "", null),
        UPPER_ROMAN("upper-roman", STNumberFormat.UPPER_ROMAN, "", null),
        THAI("thai", STNumberFormat.THAI_NUMBERS, "", null),
        /* Alphabetic */
        LOWER_ALPHA("lower-alpha", STNumberFormat.LOWER_LETTER, "", null),
        LOWER_LATIN("lower-latin", STNumberFormat.LOWER_LETTER, "", null),
        UPPER_ALPHA("upper-alpha", STNumberFormat.UPPER_LETTER, "", null),
        UPPER_LATIN("upper-latin", STNumberFormat.UPPER_LETTER, "", null),
        // hiragana
        // hiragana-iroha
        KATAKANA("katakana", STNumberFormat.AIUEO_FULL_WIDTH, "", null),
        KATAKANA_IROHA("katakana-iroha", STNumberFormat.IROHA_FULL_WIDTH, "", null),
        /* Symbolic */
        DISC("disc", STNumberFormat.BULLET, SYMBOL_DISC, FONT_WINGDINGS),
        CIRCLE("circle", STNumberFormat.BULLET, SYMBOL_CIRCLE, FONT_WINGDINGS),
        DISCLOSURE_CLOSED("disclosure-closed", STNumberFormat.BULLET, SYMBOL_DISCLOSURE_CLOSED, FONT_WINGDINGS_3),
        DISCLOSURE_OPEN("disclosure-open", STNumberFormat.BULLET, SYMBOL_DISCLOSURE_OPEN, FONT_WINGDINGS_3),
        SQUARE("square", STNumberFormat.BULLET, SYMBOL_SQUARE, FONT_WINGDINGS),
        /* Longhand East Asian */
        JAPANESE_INFORMAL("japanese-informal", STNumberFormat.JAPANESE_COUNTING, "", null),
        JAPANESE_FORMAL("japanese-formal", STNumberFormat.JAPANESE_LEGAL, "", null),
        KOREAN_HANGUL_FORMAL("korean-hangul-formal", STNumberFormat.KOREAN_COUNTING, "", null),
        // partial matching
        KOREAN_HANJA_INFORMAL("korean-hanja-informal", STNumberFormat.KOREAN_DIGITAL_2, "", null),
        // partial matching
        KOREAN_HANJA_FORMAL("korean-hanja-formal", STNumberFormat.CHINESE_LEGAL_SIMPLIFIED, "", null),
        SIMP_CHINESE_INFORMAL("simp-chinese-informal", STNumberFormat.CHINESE_COUNTING, "", null),
        SIMP_CHINESE_FORMAL("simp-chinese-formal", STNumberFormat.CHINESE_LEGAL_SIMPLIFIED, "", null),
        TRAD_CHINESE_INFORMAL("trad-chinese-informal", STNumberFormat.TAIWANESE_COUNTING, "", null),
        // partial matching
        TRAD_CHINESE_FORMAL("trad-chinese-formal", STNumberFormat.CHINESE_LEGAL_SIMPLIFIED, "", null),
        CJK_IDEOGRAPHIC("cjk-ideographic", STNumberFormat.CHINESE_LEGAL_SIMPLIFIED, "", null),
        /* ol type */
        ONE("1", STNumberFormat.DECIMAL, "", null),
        LOWER_A("a", STNumberFormat.LOWER_LETTER, "", null),
        UPPER_A("A", STNumberFormat.UPPER_LETTER, "", null),
        LOWER_I("i", STNumberFormat.LOWER_ROMAN, "", null),
        UPPER_I("I", STNumberFormat.UPPER_ROMAN, "", null),
        NONE("none", STNumberFormat.NONE, null, null);

        private static final Map<String, ListStyleType> TYPE_MAP = Arrays.stream(values())
                .collect(Collectors.toMap(Ordered::getName, Function.identity()));

        private final String name;
        private final STNumberFormat.Enum format;
        private final String text;
        private final String font;

        Ordered(String name, STNumberFormat.Enum format, String text, String font) {
            this.name = name;
            this.format = format;
            this.text = text;
            this.font = font;
        }

        @Override
        public String getName() {
            return name;
        }

        @Override
        public STNumberFormat.Enum getFormat() {
            return format;
        }

        @Override
        public String getText() {
            return text;
        }

        @Override
        public String getFont() {
            return font;
        }

        public static ListStyleType of(String type) {
            return TYPE_MAP.getOrDefault(type, DECIMAL);
        }
    }
}

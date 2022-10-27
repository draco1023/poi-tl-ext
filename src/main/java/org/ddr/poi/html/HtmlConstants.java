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

import org.apache.commons.compress.utils.Sets;
import org.ddr.poi.html.util.Colors;

import java.util.Set;

/**
 * HTML常量
 *
 * @author Draco
 * @since 2021-02-23
 */
public interface HtmlConstants {
    String TAG_A = "a";
    String TAG_IMG = "img";
    String TAG_BR = "br";
    String TAG_MATH = "math";
    String TAG_HR = "hr";
    String TAG_OL = "ol";
    String TAG_UL = "ul";
    String TAG_LI = "li";
    String TAG_TABLE = "table";
    String TAG_S = "s";
    String TAG_DEL = "del";
    /**
     * HTML5不支持strike
     */
    String TAG_STRIKE = "strike";
    String TAG_I = "i";
    String TAG_EM = "em";
    String TAG_B = "b";
    String TAG_STRONG = "strong";
    String TAG_U = "u";
    String TAG_MARK = "mark";
    String TAG_SUB = "sub";
    String TAG_SUP = "sup";
    String TAG_H1 = "h1";
    String TAG_H2 = "h2";
    String TAG_H3 = "h3";
    String TAG_H4 = "h4";
    String TAG_H5 = "h5";
    String TAG_H6 = "h6";
    /**
     * HTML5不支持big
     */
    String TAG_BIG = "big";
    String TAG_SMALL = "small";
    String TAG_CAPTION = "caption";
    String TAG_COLGROUP = "colgroup";
    String TAG_COL = "col";
    String TAG_TR = "tr";
    String TAG_TH = "th";
    String TAG_TD = "td";
    String TAG_THEAD = "thead";
    String TAG_TBODY = "tbody";
    String TAG_TFOOT = "tfoot";

    String TAG_FRAME = "frame";
    String TAG_FRAMESET = "frameset";
    String TAG_IFRAME = "iframe";
    String TAG_NOFRAMES = "noframes";
    String TAG_HTML = "html";
    String TAG_HEAD = "head";
    String TAG_BODY = "body";
    String TAG_SCRIPT = "script";
    String TAG_NOSCRIPT = "noscript";
    String TAG_TEMPLATE = "template";

    String TAG_SVG = "svg";
    String TAG_RUBY = "ruby";
    String TAG_RP = "rp";
    String TAG_RT = "rt";

    String ATTR_STYLE = "style";
    String ATTR_SRC = "src";
    String ATTR_WIDTH = "width";
    String ATTR_HEIGHT = "height";
    String ATTR_SPAN = "span";
    String ATTR_ROWSPAN = "rowspan";
    String ATTR_COLSPAN = "colspan";
    String ATTR_HREF = "href";
    String ATTR_TYPE = "type";
    /**
     * 自定义属性：行索引
     */
    String ATTR_ROW_INDEX = "_r";
    /**
     * 自定义属性：列索引
     */
    String ATTR_COLUMN_INDEX = "_c";

    String CSS_BACKGROUND = "background";
    String CSS_BACKGROUND_COLOR = "background-color";
    String CSS_BORDER = "border";
    String CSS_BORDER_STYLE = "border-style";
    String CSS_BORDER_WIDTH = "border-width";
    String CSS_BORDER_COLOR = "border-color";
    String CSS_FONT = "font";
    String CSS_MARGIN = "margin";
    String CSS_MARGIN_TOP = "margin-top";
    String CSS_MARGIN_RIGHT = "margin-right";
    String CSS_MARGIN_BOTTOM = "margin-bottom";
    String CSS_MARGIN_LEFT = "margin-left";
    String CSS_PADDING = "padding";
    String CSS_PADDING_TOP = "padding-top";
    String CSS_PADDING_RIGHT = "padding-right";
    String CSS_PADDING_BOTTOM = "padding-bottom";
    String CSS_PADDING_LEFT = "padding-left";
    String CSS_FONT_STYLE = "font-style";
    String CSS_FONT_VARIANT_CAPS = "font-variant-caps";
    String CSS_FONT_WEIGHT = "font-weight";
    String CSS_FONT_SIZE = "font-size";
    String CSS_LINE_HEIGHT = "line-height";
    String CSS_FONT_FAMILY = "font-family";
    String CSS_TEXT_DECORATION = "text-decoration";
    String CSS_TEXT_DECORATION_LINE = "text-decoration-line";
    String CSS_TEXT_DECORATION_STYLE = "text-decoration-style";
    String CSS_TEXT_DECORATION_COLOR = "text-decoration-color";
    String CSS_TEXT_INDENT = "text-indent";
    String CSS_VERTICAL_ALIGN = "vertical-align";
    String CSS_VISIBILITY = "visibility";
    String CSS_DISPLAY = "display";
    String CSS_COLOR = "color";
    String CSS_WIDTH = ATTR_WIDTH;
    String CSS_MAX_WIDTH = "max-width";
    String CSS_HEIGHT = ATTR_HEIGHT;
    String CSS_MAX_HEIGHT = "max-height";
    String CSS_BORDER_TOP = "border-top";
    String CSS_BORDER_RIGHT = "border-right";
    String CSS_BORDER_BOTTOM = "border-bottom";
    String CSS_BORDER_LEFT = "border-left";
    String CSS_BORDER_TOP_STYLE = "border-top-style";
    String CSS_BORDER_RIGHT_STYLE = "border-right-style";
    String CSS_BORDER_BOTTOM_STYLE = "border-bottom-style";
    String CSS_BORDER_LEFT_STYLE = "border-left-style";
    String CSS_BORDER_TOP_WIDTH = "border-top-width";
    String CSS_BORDER_RIGHT_WIDTH = "border-right-width";
    String CSS_BORDER_BOTTOM_WIDTH = "border-bottom-width";
    String CSS_BORDER_LEFT_WIDTH = "border-left-width";
    String CSS_BORDER_TOP_COLOR = "border-top-color";
    String CSS_BORDER_RIGHT_COLOR = "border-right-color";
    String CSS_BORDER_BOTTOM_COLOR = "border-bottom-color";
    String CSS_BORDER_LEFT_COLOR = "border-left-color";
    String CSS_FLOAT = "float";
    String CSS_WHITE_SPACE = "white-space";
    String CSS_LIST_STYLE = "list-style";
    String CSS_LIST_STYLE_TYPE = "list-style-type";
    String CSS_BORDER_COLLAPSE = "border-collapse";
    String CSS_BORDER_SPACING = "border-spacing";
    String CSS_CAPTION_SIDE = "caption-side";
    String CSS_LETTER_SPACING = "letter-spacing";
    String CSS_TEXT_ALIGN = "text-align";

    String NORMAL = "normal";
    String ITALIC = "italic";
    String OBLIQUE = "oblique";
    String SMALL_CAPS = "small-caps";
    String BOLD = "bold";
    String BOLDER = "bolder";
    String LIGHTER = "lighter";
    String START = "start";
    String LEFT = "left";
    String END = "end";
    String RIGHT = "right";
    String CENTER = "center";
    String JUSTIFY = "justify";
    String JUSTIFY_ALL = "justify-all";
    String TOP = "top";
    String BOTTOM = "bottom";
    String MIDDLE = "middle";

    String XX_SMALL = "xx-small";
    String X_SMALL = "x-small";
    String SMALL = "small";
    String MEDIUM = "medium";
    String LARGE = "large";
    String X_LARGE = "x-large";
    String XX_LARGE = "xx-large";
    String XXX_LARGE = "xxx-large";

    String SMALLER = "smaller";
    String LARGER = "larger";

    String THIN = "thin";
    String THICK = "thick";

    String PT = "pt";
    String PC = "pc";
    String IN = "in";
    String CM = "cm";
    String MM = "mm";
    String PX = "px";
    String EM = "em";
    String REM = "rem";
    String VW = "vw";
    String VH = "vh";
    String VMIN = "vmin";
    String VMAX = "vmax";
    String PERCENT = "%";
    // 自定义单位
    String EMU = "emu";
    /**
     * dxa的单位，twentieth of a point = 1 / 20 pt
     */
    String TWIP = "twip";

    String SLASH = "/";
    String COMMA = ",";
    String COLON = ":";
    String SHARP = "#";
    String SEMICOLON = ";";
    String QUESTION = "?";
    String PLUS = "+";
    String MINUS = "-";

    String LINE_THROUGH = "line-through";
    String UNDERLINE = "underline";
    String SOLID = "solid";
    String DOUBLE = "double";
    String DOTTED = "dotted";
    String DASHED = "dashed";
    String WAVY = "wavy";
    String NONE = "none";

    String GROOVE = "groove";
    String RIDGE = "ridge";
    // 类似groove
    String INSET = "inset";
    // 类似ridge
    String OUTSET = "outset";

    String HIDDEN = "hidden";
    String COLLAPSE = "collapse";

    String SUPER = "super";
    String SUB = "sub";

    String NO_WRAP = "nowrap";
    String PRE = "pre";
    String PRE_WRAP = "pre-wrap";
    String PRE_LINE = "pre-line";
    String BREAK_SPACES = "break-spaces";

    Set<String> FONT_STYLES = Sets.newHashSet(NORMAL, ITALIC, OBLIQUE);
    Set<String> FONT_VARIANTS = Sets.newHashSet(NORMAL, SMALL_CAPS);
    Set<String> FONT_WEIGHTS = Sets.newHashSet(NORMAL, BOLD, BOLDER, LIGHTER);
    Set<String> BORDER_STYLES = Sets.newHashSet(NONE, HIDDEN, DOTTED, DASHED, SOLID, DOUBLE, GROOVE, RIDGE, INSET, OUTSET);
    // 不支持overline
    Set<String> TEXT_DECORATION_LINES = Sets.newHashSet(UNDERLINE, LINE_THROUGH);
    Set<String> TEXT_DECORATION_STYLES = Sets.newHashSet(SOLID, DOUBLE, DOTTED, DASHED, WAVY);

    /**
     * 可继承的样式
     * <a href="https://www.w3.org/TR/CSS22/propidx.html">Specification</a>
     */
    Set<String> INHERITABLE_STYLES = Sets.newHashSet(
            "azimuth",
            CSS_BORDER_COLLAPSE,
            CSS_BORDER_SPACING,
            CSS_CAPTION_SIDE,
            CSS_COLOR,
            "cursor",
            "direction",
            "elevation",
            "empty-cells",
            CSS_FONT_FAMILY,
            CSS_FONT_SIZE,
            CSS_FONT_STYLE,
            CSS_FONT_VARIANT_CAPS,
            CSS_FONT_WEIGHT,
            CSS_FONT,
            CSS_LETTER_SPACING,
            CSS_LINE_HEIGHT,
            "list-style-image",
            "list-style-position",
            CSS_LIST_STYLE_TYPE,
            CSS_LIST_STYLE,
            "orphans",
            "pitch-range",
            "pitch",
            "quotes",
            "richness",
            "speak-header",
            "speak-numeral",
            "speak-punctuation",
            "speak",
            "speech-rate",
            "stress",
            CSS_TEXT_ALIGN,
            CSS_TEXT_INDENT,
            "text-transform",
            CSS_VISIBILITY,
            "voice-family",
            "volume",
            CSS_WHITE_SPACE,
            "widows",
            "word-spacing"
    );

    /**
     * 需要保留的空标签
     */
    Set<String> KEEP_EMPTY_TAGS = Sets.newHashSet(TAG_LI, TAG_HR);

    /**
     * Word中一些主要的默认字体
     */
    Set<String> MAJOR_FONT = Sets.newHashSet("宋体", "SIMSUN", "新細明體", "TIMES NEW ROMAN", "ARIAL");

    String DEFINED_ITALIC = inlineStyle(CSS_FONT_STYLE, ITALIC);
    String DEFINED_STRIKE = inlineStyle(CSS_TEXT_DECORATION_LINE, LINE_THROUGH);
    String DEFINED_BOLD = inlineStyle(CSS_FONT_WEIGHT, BOLD);
    String DEFINED_UNDERLINE = inlineStyle(CSS_TEXT_DECORATION_LINE, UNDERLINE);
    String DEFINED_SUPERSCRIPT = inlineStyle(CSS_VERTICAL_ALIGN, SUPER);
    String DEFINED_SUBSCRIPT = inlineStyle(CSS_VERTICAL_ALIGN, SUB);
    String DEFINED_LARGER = inlineStyle(CSS_FONT_SIZE, LARGER);
    String DEFINED_SMALLER = inlineStyle(CSS_FONT_SIZE, SMALLER);
    String DEFINED_MARK = inlineStyle(CSS_BACKGROUND_COLOR, Colors.getColorByName("yellow"));

    /**
     * 生成行内样式声明
     *
     * @param key 样式属性
     * @param value 样式值
     * @return 行内样式声明
     */
    static String inlineStyle(String key, String value) {
        return key + COLON + value + SEMICOLON;
    }

    /**
     * @param fontName 字体名称
     * @return 是否为主要字体
     */
    static boolean isMajorFont(String fontName) {
        return MAJOR_FONT.contains(fontName.toUpperCase());
    }
}

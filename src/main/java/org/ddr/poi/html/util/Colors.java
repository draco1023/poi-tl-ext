package org.ddr.poi.html.util;

import org.apache.commons.lang3.StringUtils;
import org.ddr.poi.html.HtmlConstants;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashMap;
import java.util.Map;

/**
 * 颜色工具类
 *
 * @author Draco
 * @since 2021-02-23
 */
public class Colors {
    private static final Logger log = LoggerFactory.getLogger(Colors.class);

    private static final Map<String, String> COLOR_MAP = new HashMap<>(160);

    public static final String BLACK = "000000";
    public static final String WHITE = "FFFFFF";
    public static final String DEFAULT_COLOR = BLACK;

    // https://developer.mozilla.org/zh-CN/docs/Web/CSS/color_value#%E8%89%B2%E5%BD%A9%E5%85%B3%E9%94%AE%E5%AD%97
    static {
        COLOR_MAP.put("black", BLACK);
        COLOR_MAP.put("silver", "C0C0C0");
        COLOR_MAP.put("gray", "808080");
        COLOR_MAP.put("white", "FFFFFF");
        COLOR_MAP.put("maroon", "800000");
        COLOR_MAP.put("red", "FF0000");
        COLOR_MAP.put("purple", "800080");
        COLOR_MAP.put("fuchsia", "FF00FF");
        COLOR_MAP.put("green", "008000");
        COLOR_MAP.put("lime", "00FF00");
        COLOR_MAP.put("olive", "808000");
        COLOR_MAP.put("yellow", "FFFF00");
        COLOR_MAP.put("navy", "000080");
        COLOR_MAP.put("blue", "0000FF");
        COLOR_MAP.put("teal", "008080");
        COLOR_MAP.put("aqua", "00FFFF");
        COLOR_MAP.put("orange", "FFA500");
        COLOR_MAP.put("aliceblue", "F0F8FF");
        COLOR_MAP.put("antiquewhite", "FAEBD7");
        COLOR_MAP.put("aquamarine", "7FFFD4");
        COLOR_MAP.put("azure", "F0FFFF");
        COLOR_MAP.put("beige", "F5F5DC");
        COLOR_MAP.put("bisque", "FFE4C4");
        COLOR_MAP.put("blanchedalmond", "FFEBCD");
        COLOR_MAP.put("blueviolet", "8A2BE2");
        COLOR_MAP.put("brown", "A52A2A");
        COLOR_MAP.put("burlywood", "DEB887");
        COLOR_MAP.put("cadetblue", "5F9EA0");
        COLOR_MAP.put("chartreuse", "7FFF00");
        COLOR_MAP.put("chocolate", "D2691E");
        COLOR_MAP.put("coral", "FF7F50");
        COLOR_MAP.put("cornflowerblue", "6495ED");
        COLOR_MAP.put("cornsilk", "FFF8DC");
        COLOR_MAP.put("crimson", "DC143C");
        COLOR_MAP.put("darkblue", "00008B");
        COLOR_MAP.put("darkcyan", "008B8B");
        COLOR_MAP.put("darkgoldenrod", "B8860B");
        COLOR_MAP.put("darkgray", "A9A9A9");
        COLOR_MAP.put("darkgreen", "006400");
        COLOR_MAP.put("darkgrey", "A9A9A9");
        COLOR_MAP.put("darkkhaki", "BDB76B");
        COLOR_MAP.put("darkmagenta", "8B008B");
        COLOR_MAP.put("darkolivegreen", "556B2F");
        COLOR_MAP.put("darkorange", "FF8C00");
        COLOR_MAP.put("darkorchid", "9932CC");
        COLOR_MAP.put("darkred", "8B0000");
        COLOR_MAP.put("darksalmon", "E9967A");
        COLOR_MAP.put("darkseagreen", "8FBC8F");
        COLOR_MAP.put("darkslateblue", "483D8B");
        COLOR_MAP.put("darkslategray", "2F4F4F");
        COLOR_MAP.put("darkslategrey", "2F4F4F");
        COLOR_MAP.put("darkturquoise", "00CED1");
        COLOR_MAP.put("darkviolet", "9400D3");
        COLOR_MAP.put("deeppink", "FF1493");
        COLOR_MAP.put("deepskyblue", "00BFFF");
        COLOR_MAP.put("dimgray", "696969");
        COLOR_MAP.put("dimgrey", "696969");
        COLOR_MAP.put("dodgerblue", "1E90FF");
        COLOR_MAP.put("firebrick", "B22222");
        COLOR_MAP.put("floralwhite", "FFFAF0");
        COLOR_MAP.put("forestgreen", "228B22");
        COLOR_MAP.put("gainsboro", "DCDCDC");
        COLOR_MAP.put("ghostwhite", "F8F8FF");
        COLOR_MAP.put("gold", "FFD700");
        COLOR_MAP.put("goldenrod", "DAA520");
        COLOR_MAP.put("greenyellow", "ADFF2F");
        COLOR_MAP.put("grey", "808080");
        COLOR_MAP.put("honeydew", "F0FFF0");
        COLOR_MAP.put("hotpink", "FF69B4");
        COLOR_MAP.put("indianred", "CD5C5C");
        COLOR_MAP.put("indigo", "4B0082");
        COLOR_MAP.put("ivory", "FFFFF0");
        COLOR_MAP.put("khaki", "F0E68C");
        COLOR_MAP.put("lavender", "E6E6FA");
        COLOR_MAP.put("lavenderblush", "FFF0F5");
        COLOR_MAP.put("lawngreen", "7CFC00");
        COLOR_MAP.put("lemonchiffon", "FFFACD");
        COLOR_MAP.put("lightblue", "ADD8E6");
        COLOR_MAP.put("lightcoral", "F08080");
        COLOR_MAP.put("lightcyan", "E0FFFF");
        COLOR_MAP.put("lightgoldenrodyellow", "FAFAD2");
        COLOR_MAP.put("lightgray", "D3D3D3");
        COLOR_MAP.put("lightgreen", "90EE90");
        COLOR_MAP.put("lightgrey", "D3D3D3");
        COLOR_MAP.put("lightpink", "FFB6C1");
        COLOR_MAP.put("lightsalmon", "FFA07A");
        COLOR_MAP.put("lightseagreen", "20B2AA");
        COLOR_MAP.put("lightskyblue", "87CEFA");
        COLOR_MAP.put("lightslategray", "778899");
        COLOR_MAP.put("lightslategrey", "778899");
        COLOR_MAP.put("lightsteelblue", "B0C4DE");
        COLOR_MAP.put("lightyellow", "FFFFE0");
        COLOR_MAP.put("limegreen", "32CD32");
        COLOR_MAP.put("linen", "FAF0E6");
        COLOR_MAP.put("mediumaquamarine", "66CDAA");
        COLOR_MAP.put("mediumblue", "0000CD");
        COLOR_MAP.put("mediumorchid", "BA55D3");
        COLOR_MAP.put("mediumpurple", "9370DB");
        COLOR_MAP.put("mediumseagreen", "3CB371");
        COLOR_MAP.put("mediumslateblue", "7B68EE");
        COLOR_MAP.put("mediumspringgreen", "00FA9A");
        COLOR_MAP.put("mediumturquoise", "48D1CC");
        COLOR_MAP.put("mediumvioletred", "C71585");
        COLOR_MAP.put("midnightblue", "191970");
        COLOR_MAP.put("mintcream", "F5FFFA");
        COLOR_MAP.put("mistyrose", "FFE4E1");
        COLOR_MAP.put("moccasin", "FFE4B5");
        COLOR_MAP.put("navajowhite", "FFDEAD");
        COLOR_MAP.put("oldlace", "FDF5E6");
        COLOR_MAP.put("olivedrab", "6B8E23");
        COLOR_MAP.put("orangered", "FF4500");
        COLOR_MAP.put("orchid", "DA70D6");
        COLOR_MAP.put("palegoldenrod", "EEE8AA");
        COLOR_MAP.put("palegreen", "98FB98");
        COLOR_MAP.put("paleturquoise", "AFEEEE");
        COLOR_MAP.put("palevioletred", "DB7093");
        COLOR_MAP.put("papayawhip", "FFEFD5");
        COLOR_MAP.put("peachpuff", "FFDAB9");
        COLOR_MAP.put("peru", "CD853F");
        COLOR_MAP.put("pink", "FFC0CB");
        COLOR_MAP.put("plum", "DDA0DD");
        COLOR_MAP.put("powderblue", "B0E0E6");
        COLOR_MAP.put("rosybrown", "BC8F8F");
        COLOR_MAP.put("royalblue", "4169E1");
        COLOR_MAP.put("saddlebrown", "8B4513");
        COLOR_MAP.put("salmon", "FA8072");
        COLOR_MAP.put("sandybrown", "F4A460");
        COLOR_MAP.put("seagreen", "2E8B57");
        COLOR_MAP.put("seashell", "FFF5EE");
        COLOR_MAP.put("sienna", "A0522D");
        COLOR_MAP.put("skyblue", "87CEEB");
        COLOR_MAP.put("slateblue", "6A5ACD");
        COLOR_MAP.put("slategray", "708090");
        COLOR_MAP.put("slategrey", "708090");
        COLOR_MAP.put("snow", "FFFAFA");
        COLOR_MAP.put("springgreen", "00FF7F");
        COLOR_MAP.put("steelblue", "4682B4");
        COLOR_MAP.put("tan", "D2B48C");
        COLOR_MAP.put("thistle", "D8BFD8");
        COLOR_MAP.put("tomato", "FF6347");
        COLOR_MAP.put("turquoise", "40E0D0");
        COLOR_MAP.put("violet", "EE82EE");
        COLOR_MAP.put("wheat", "F5DEB3");
        COLOR_MAP.put("whitesmoke", "F5F5F5");
        COLOR_MAP.put("yellowgreen", "9ACD32");
        COLOR_MAP.put("rebeccapurple", "663399");
    }

    /**
     * 根据颜色名称获取颜色值，未找到时返回null
     *
     * @param name 颜色名称
     * @return 颜色值
     */
    public static String getColorByName(String name) {
        return name == null ? null : COLOR_MAP.get(name.toLowerCase());
    }

    /**
     * 根据颜色名称获取颜色值，未找到时返回默认颜色值
     *
     * @param name 颜色名称
     * @param defaultColor 默认颜色值
     * @return 颜色值
     */
    public static String getColorByName(String name, String defaultColor) {
        return name == null ? defaultColor : COLOR_MAP.getOrDefault(name.toLowerCase(), defaultColor);
    }

    /**
     * hsl颜色转换为颜色值
     *
     * @param h hue 0 - 360
     * @param s saturation 1 - 100
     * @param l lightness 1 - 100
     * @return 颜色值
     */
    public static String fromHSL(float h, float s, float l) {
        float q, p, r, g, b;

        if (s == 0) {
            r = g = b = l; // achromatic
        } else {
            q = l < 0.5 ? (l * (1 + s)) : (l + s - l * s);
            p = 2 * l - q;
            r = hue2rgb(p, q, h + 1.0f / 3);
            g = hue2rgb(p, q, h);
            b = hue2rgb(p, q, h - 1.0f / 3);
        }
        return toHexString(toInt(r), toInt(g), toInt(b));
    }

    private static float hue2rgb(float p, float q, float h) {
        if (h < 0) {
            h += 1;
        }

        if (h > 1) {
            h -= 1;
        }

        if (6 * h < 1) {
            return p + ((q - p) * 6 * h);
        }

        if (2 * h < 1) {
            return q;
        }

        if (3 * h < 2) {
            return p + ((q - p) * 6 * ((2.0f / 3.0f) - h));
        }

        return p;
    }

    private static int toInt(float f) {
        return (int) (f * 255 + 0.5);
    }

    /**
     * 将RGB转换为颜色值
     *
     * @param r Red
     * @param g Green
     * @param b Blue
     * @return 颜色值
     */
    public static String toHexString(int r, int g, int b) {
        return String.format("%02X%02X%02X", r, g, b);
    }

    /**
     * 解析样式值为颜色值
     *
     * @param style 样式值
     * @param defaultColor 默认颜色值
     * @return 颜色值
     */
    public static String fromStyle(String style, String defaultColor) {
        if (StringUtils.isBlank(style)) {
            return defaultColor;
        }
        // Word中不支持alpha通道，直接忽略
        if (style.startsWith(HtmlConstants.SHARP)) {
            String hex = style.substring(1).trim();
            if (hex.length() == 3 || hex.length() == 4) {
                char[] chars = new char[6];
                for (int i = 0; i < 6; i++) {
                    chars[i] = hex.charAt(i >> 1);
                }
                return String.valueOf(chars);
            } else if (hex.length() >= 6) {
                return hex.substring(0, 6);
            } else {
                warn(style);
            }
        } else if (style.startsWith("rgb")) {
//            color: rgb(34, 12, 64, 0.6);
//            color: rgba(34, 12, 64, 0.6);
//            color: rgb(34 12 64 / 0.6);
//            color: rgba(34 12 64 / 0.3);
//            color: rgb(34.0 12 64 / 60%);
//            color: rgba(34.6 12 64 / 30%);
            String[] array = StringUtils.split(StringUtils.substringBetween(style, "(", ")"), ", /");
            if (array != null && array.length >= 3) {
                return toHexString((int) Float.parseFloat(array[0]), (int) Float.parseFloat(array[1]), (int) Float.parseFloat(array[2]));
            } else {
                warn(style);
            }
        } else if (style.startsWith("hsl")) {
//            color: hsl(30, 100%, 50%, 0.6);
//            color: hsla(30, 100%, 50%, 0.6);
//            color: hsl(30 100% 50% / 0.6);
//            color: hsla(30 100% 50% / 0.6);
//            color: hsl(30.0 100% 50% / 60%);
//            color: hsla(30.2 100% 50% / 60%);
            String[] array = StringUtils.split(StringUtils.substringBetween(style, "(", ")"), ", /");
            if (array != null && array.length >= 3
                    && array[1].endsWith(HtmlConstants.PERCENT) && array[2].endsWith(HtmlConstants.PERCENT)) {
                float h = Float.parseFloat(array[0]);
                String ss = array[1];
                float s = Float.parseFloat(ss.substring(0, ss.length() - 1)) / 100;
                String ls = array[2];
                float l = Float.parseFloat(ls.substring(0, ls.length() - 1)) / 100;
                return fromHSL(h, s, l);
            } else {
                warn(style);
            }
        } else {
            return getColorByName(style, defaultColor);
        }

        return defaultColor;
    }

    /**
     * 解析样式值为颜色值，解析失败时返回默认颜色值（黑色）
     *
     * @param style 样式值
     * @return 颜色值
     */
    public static String fromStyle(String style) {
        return fromStyle(style, DEFAULT_COLOR);
    }

    private static void warn(String style) {
        log.warn("Illegal color: {}", style);
    }
}

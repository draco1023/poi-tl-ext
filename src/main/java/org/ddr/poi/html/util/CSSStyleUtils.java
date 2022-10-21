package org.ddr.poi.html.util;

import com.steadystate.css.dom.CSSStyleDeclarationImpl;
import com.steadystate.css.dom.CSSValueImpl;
import com.steadystate.css.dom.Property;
import com.steadystate.css.parser.CSSOMParser;
import com.steadystate.css.parser.SACParserCSS3;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.ddr.poi.html.HtmlConstants;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.css.sac.InputSource;
import org.w3c.dom.css.CSSValue;

import java.io.IOException;
import java.io.StringReader;

/**
 * CSS样式相关的工具类
 *
 * @author Draco
 * @since 2021-10-11
 */
public class CSSStyleUtils {
    private static final Logger log = LoggerFactory.getLogger(CSSStyleUtils.class);

    /**
     * CSS解析器
     */
    public static final CSSOMParser CSS_PARSER = new CSSOMParser(new SACParserCSS3());
    /**
     * 空样式
     */
    public static final CSSStyleDeclarationImpl EMPTY_STYLE = new EmptyCSSStyle();

    /**
     * 样式是否为空
     *
     * @param style 样式声明
     * @return 是否为空
     */
    public static boolean isEmpty(CSSStyleDeclarationImpl style) {
        return style == null || EMPTY_STYLE.equals(style) || style.getProperties().isEmpty();
    }

    /**
     * 解析行内样式，解析失败时返回默认空样式
     *
     * @param inlineStyle 行内样式声明
     * @return 样式
     */
    public static CSSStyleDeclarationImpl parse(String inlineStyle) {
        try (StringReader sr = new StringReader(inlineStyle)) {
            return (CSSStyleDeclarationImpl) CSS_PARSER.parseStyleDeclaration(new InputSource(sr));
        } catch (IOException e) {
            log.warn("Inline style parse error: {}", inlineStyle, e);
            return EMPTY_STYLE;
        }
    }

    /**
     * 解析行内样式，解析失败时返回新的空样式实例
     *
     * @param inlineStyle 行内样式声明
     * @return 样式
     */
    public static CSSStyleDeclarationImpl parseNew(String inlineStyle) {
        try (StringReader sr = new StringReader(inlineStyle)) {
            return (CSSStyleDeclarationImpl) CSS_PARSER.parseStyleDeclaration(new InputSource(sr));
        } catch (IOException e) {
            log.warn("Inline style parse error: {}", inlineStyle, e);
            return new CSSStyleDeclarationImpl();
        }
    }

    /**
     * 解析样式值
     *
     * @param value 样式值字符串
     * @return 样式值
     */
    public static CSSValue parseValue(String value) {
        try (StringReader sr = new StringReader(value)) {
            return CSS_PARSER.parsePropertyValue(new InputSource(sr));
        } catch (IOException e) {
            log.warn("CSS value parse error: {}", value, e);
            return new CSSValueImpl();
        }
    }

    /**
     * 将样式键值对转换为样式属性
     *
     * @param key 样式名称
     * @param value 样式值
     * @return 样式属性
     */
    public static Property newProperty(String key, String value) {
        return new Property(key, parseValue(value), false);
    }

    /**
     * 分解缩写的样式
     *
     * @param style 样式声明
     */
    public static void split(CSSStyleDeclarationImpl style) {
        for (int i = style.getProperties().size() - 1; i >= 0; i--) {
            final Property p = style.getProperties().get(i);
            if (p != null && p.getValue() != null) {
                String name = p.getName().toLowerCase();
                CSSValueImpl valueList = (CSSValueImpl) p.getValue();
                int length = valueList.getLength();
                // 将复合样式拆分成单属性样式
                switch (name) {
                    case HtmlConstants.CSS_BACKGROUND:
                        splitBackground(valueList, length, style, i);
                        break;
                    case HtmlConstants.CSS_BORDER:
                        splitBorder(valueList, length, style, i);
                        break;
                    case HtmlConstants.CSS_BORDER_TOP:
                        splitBorder(valueList, length, style, i, HtmlConstants.CSS_BORDER_TOP_STYLE,
                                HtmlConstants.CSS_BORDER_TOP_WIDTH, HtmlConstants.CSS_BORDER_TOP_COLOR);
                        break;
                    case HtmlConstants.CSS_BORDER_RIGHT:
                        splitBorder(valueList, length, style, i, HtmlConstants.CSS_BORDER_RIGHT_STYLE,
                                HtmlConstants.CSS_BORDER_RIGHT_WIDTH, HtmlConstants.CSS_BORDER_RIGHT_COLOR);
                        break;
                    case HtmlConstants.CSS_BORDER_BOTTOM:
                        splitBorder(valueList, length, style, i, HtmlConstants.CSS_BORDER_BOTTOM_STYLE,
                                HtmlConstants.CSS_BORDER_BOTTOM_WIDTH, HtmlConstants.CSS_BORDER_BOTTOM_COLOR);
                        break;
                    case HtmlConstants.CSS_BORDER_LEFT:
                        splitBorder(valueList, length, style, i, HtmlConstants.CSS_BORDER_LEFT_STYLE,
                                HtmlConstants.CSS_BORDER_LEFT_WIDTH, HtmlConstants.CSS_BORDER_LEFT_COLOR);
                        break;
                    case HtmlConstants.CSS_BORDER_STYLE:
                        splitBox(valueList, length, style, i, BoxProperty.BORDER_STYLE);
                        break;
                    case HtmlConstants.CSS_BORDER_WIDTH:
                        splitBox(valueList, length, style, i, BoxProperty.BORDER_WIDTH);
                        break;
                    case HtmlConstants.CSS_BORDER_COLOR:
                        splitBox(valueList, length, style, i, BoxProperty.BORDER_COLOR);
                        break;
                    case HtmlConstants.CSS_FONT:
                        splitFont(valueList, length, style, i);
                        break;
                    case HtmlConstants.CSS_MARGIN:
                        splitBox(valueList, length, style, i, BoxProperty.MARGIN);
                        break;
                    case HtmlConstants.CSS_PADDING:
                        splitBox(valueList, length, style, i, BoxProperty.PADDING);
                        break;
                    case HtmlConstants.CSS_LIST_STYLE:
                        splitListStyle(valueList, length, style, i);
                        break;
                    case HtmlConstants.CSS_TEXT_DECORATION:
                        splitTextDecoration(valueList, length, style, i);
                        break;
                }
            }
        }
    }

    private static void splitBackground(CSSValueImpl valueList, int length, CSSStyleDeclarationImpl style, int i) {
        if (length == 0) {
            String cssText = valueList.getCssText().toLowerCase();
            String color = Colors.fromStyle(cssText, null);
            if (color != null) {
                style.getProperties().add(i, new Property(HtmlConstants.CSS_BACKGROUND_COLOR, valueList, false));
            }
        } else {
            for (int j = 0; j < length; j++) {
                CSSValue item = valueList.item(j);
                String cssText = item.getCssText().toLowerCase();
                String color = Colors.fromStyle(cssText, null);
                if (color != null) {
                    style.getProperties().add(i, new Property(HtmlConstants.CSS_BACKGROUND_COLOR, item, false));
                    break;
                }
            }
        }
    }

    private static void splitBorder(CSSValueImpl valueList, int length, CSSStyleDeclarationImpl style, int i) {
        if (length == 0) {
            String cssText = valueList.getCssText();
            if (StringUtils.isNotBlank(cssText)) {
                handleBorderValue(style, i, valueList, cssText);
            }
        } else {
            for (int j = 0; j < length; j++) {
                CSSValue item = valueList.item(j);
                String value = item.getCssText();
                handleBorderValue(style, i, item, value);
            }
        }
    }

    private static void splitBorder(CSSValueImpl valueList, int length, CSSStyleDeclarationImpl style, int i,
                                    String styleProperty, String widthProperty, String colorProperty) {
        if (length == 0) {
            String cssText = valueList.getCssText();
            if (StringUtils.isNotBlank(cssText)) {
                handleBorderValue(style, i, valueList, cssText, styleProperty, widthProperty, colorProperty);
            }
        } else {
            for (int j = 0; j < length; j++) {
                CSSValue item = valueList.item(j);
                String value = item.getCssText();
                handleBorderValue(style, i, item, value, styleProperty, widthProperty, colorProperty);
            }
        }
    }

    private static void handleBorderValue(CSSStyleDeclarationImpl style, int i, CSSValue item, String value) {
        value = value.toLowerCase();
        if (HtmlConstants.BORDER_STYLES.contains(value)) {
            BoxProperty.BORDER_STYLE.setValues(style, i, item);
        } else if (NamedBorderWidth.contains(value)) {
            BoxProperty.BORDER_WIDTH.setValues(style, i, item);
        } else if (Character.isDigit(value.charAt(0))) {
            CSSLength width = CSSLength.of(value);
            if (width.isValid()) {
                BoxProperty.BORDER_WIDTH.setValues(style, i, item);
            }
        } else {
            BoxProperty.BORDER_COLOR.setValues(style, i, item);
        }
    }

    private static void handleBorderValue(CSSStyleDeclarationImpl style, int i, CSSValue item, String value,
                                          String styleProperty, String widthProperty, String colorProperty) {
        value = value.toLowerCase();
        if (HtmlConstants.BORDER_STYLES.contains(value)) {
            style.getProperties().add(i, new Property(styleProperty, item, false));
        } else if (NamedBorderWidth.contains(value)) {
            style.getProperties().add(i, new Property(widthProperty, item, false));
        } else if (Character.isDigit(value.charAt(0))) {
            CSSLength width = CSSLength.of(value);
            if (width.isValid()) {
                style.getProperties().add(i, new Property(widthProperty, item, false));
            }
        } else {
            style.getProperties().add(i, new Property(colorProperty, item, false));
        }
    }

    private static void splitFont(CSSValueImpl valueList, int length, CSSStyleDeclarationImpl style, int i) {
        if (length == 0) {
            return;
        }

        boolean styleHandled = false;
        boolean sizeHandled = false;
        for (int j = 0; j < length; j++) {
            CSSValue item = valueList.item(j);
            String value = item.getCssText();
            String lowerCase = value.toLowerCase();
            if (!styleHandled && HtmlConstants.FONT_STYLES.contains(lowerCase)) {
                style.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_STYLE, item, false));
                styleHandled = true;
            } else if (HtmlConstants.FONT_VARIANTS.contains(lowerCase)) {
                style.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_VARIANT_CAPS, item, false));
            } else if (HtmlConstants.FONT_WEIGHTS.contains(lowerCase) || NumberUtils.isParsable(value)) {
                style.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_WEIGHT, item, false));
            } else if (HtmlConstants.SLASH.equals(value)) {
                // 字号与行高分隔符
                // https://www.w3.org/TR/CSS22/fonts.html#value-def-absolute-size
                // xx-small, x-small, small, medium, large, x-large, xx-large, xxx-large
                // 1,        ,        2,     3,      4,     5,       6,        7
                // FIXME font元素由于已废弃暂不支持
                // 长度/百分比
                CSSValue fontSize = valueList.item(j - 1);
                style.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_SIZE, fontSize, false));
                sizeHandled = true;
                if (++j < length) {
                    // 数字/长度/百分比
                    CSSValue lineHeight = valueList.item(j);
                    style.getProperties().add(i, new Property(HtmlConstants.CSS_LINE_HEIGHT, lineHeight, false));
                }
            } else if (HtmlConstants.COMMA.equals(value)) {
                // 多个字体之间的分隔符
                CSSValue firstFont = valueList.item(j - 1);
                if (!sizeHandled) {
                    CSSValue fontSize = valueList.item(j - 2);
                    style.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_SIZE, fontSize, false));
                }
                if (HtmlConstants.isMajorFont(firstFont.getCssText())) {
                    style.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_FAMILY, firstFont, false));
                } else {
                    for (j++; j < length; j++) {
                        CSSValue fontFamily = valueList.item(j);
                        if (HtmlConstants.isMajorFont(fontFamily.getCssText())) {
                            style.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_FAMILY, fontFamily, false));
                            break;
                        }
                    }
                }
                break;
            } else if (j == length - 1) {
                // font-family在font中一定是最后出现
                style.getProperties().add(i, new Property(HtmlConstants.CSS_FONT_FAMILY, item, false));
            }
        }
    }

    private static void splitBox(CSSValueImpl valueList, int length, CSSStyleDeclarationImpl style, int i,
                                 BoxProperty boxProperty) {
        switch (length) {
            // 当仅一个值时实际返回长度为0
            case 0:
            case 1:
                if (StringUtils.isNotBlank(valueList.getCssText())) {
                    boxProperty.setValues(style, i, valueList);
                }
                break;
            case 2:
                boxProperty.setValues(style, i, valueList.item(0), valueList.item(1));
                break;
            case 3:
                boxProperty.setValues(style, i, valueList.item(0), valueList.item(1), valueList.item(2));
                break;
            case 4:
                boxProperty.setValues(style, i, valueList.item(0), valueList.item(1), valueList.item(2), valueList.item(3));
                break;
        }
    }

    private static void splitListStyle(CSSValueImpl valueList, int length, CSSStyleDeclarationImpl style, int i) {
        switch (length) {
            case 0:
            case 1:
                if (StringUtils.isNotBlank(valueList.getCssText())) {
                    style.getProperties().add(i, new Property(HtmlConstants.CSS_LIST_STYLE_TYPE, valueList, false));
                }
                break;
            default:
                style.getProperties().add(i, new Property(HtmlConstants.CSS_LIST_STYLE_TYPE, valueList.item(0), false));
        }
    }

    private static void splitTextDecoration(CSSValueImpl valueList, int length, CSSStyleDeclarationImpl style, int i) {
        if (length == 0) {
            style.getProperties().add(i, new Property(HtmlConstants.CSS_TEXT_DECORATION_LINE, valueList, false));
            return;
        }

        StringBuilder lines = new StringBuilder(22);
        for (int j = 0; j < length; j++) {
            CSSValue item = valueList.item(j);
            String value = item.getCssText();
            String lowerCase = value.toLowerCase();
            if (HtmlConstants.NONE.equals(lowerCase)) {
                style.getProperties().add(i, new Property(HtmlConstants.CSS_TEXT_DECORATION_LINE, item, false));
                break;
            }
            if (HtmlConstants.TEXT_DECORATION_LINES.contains(lowerCase)) {
                if (lines.length() > 0) {
                    lines.append(' ');
                }
                lines.append(lowerCase);
            } else if (HtmlConstants.TEXT_DECORATION_STYLES.contains(lowerCase)) {
                style.getProperties().add(i, new Property(HtmlConstants.CSS_TEXT_DECORATION_STYLE, item, false));
            } else if (Colors.maybe(lowerCase)) {
                style.getProperties().add(i, new Property(HtmlConstants.CSS_TEXT_DECORATION_COLOR, item, false));
            }
        }

        if (lines.length() > 0) {
            style.getProperties().add(i, newProperty(HtmlConstants.CSS_TEXT_DECORATION_LINE, lines.toString()));
        }
    }
}

package org.ddr.poi.html.util;

import com.steadystate.css.dom.CSSStyleDeclarationImpl;

/**
 * 表格列样式定义
 *
 * @author Draco
 * @since 2022-10-28
 */
public class ColumnStyle {
    private CSSStyleDeclarationImpl style;
    private CSSLength width;

    public ColumnStyle(CSSStyleDeclarationImpl style, CSSLength width) {
        this.style = style;
        this.width = width;
    }

    public CSSStyleDeclarationImpl getStyle() {
        return style;
    }

    public CSSLength getWidth() {
        return width;
    }
}

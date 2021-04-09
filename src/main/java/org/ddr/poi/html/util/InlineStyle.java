package org.ddr.poi.html.util;

import com.steadystate.css.dom.CSSStyleDeclarationImpl;

/**
 * 行内样式封装类
 *
 * @author Draco
 * @since 2021-03-18
 */
public final class InlineStyle {
    /**
     * 样式声明
     */
    private final CSSStyleDeclarationImpl declaration;
    /**
     * 是否为区块元素
     */
    private final boolean block;

    public InlineStyle(CSSStyleDeclarationImpl declaration, boolean block) {
        this.declaration = declaration;
        this.block = block;
    }

    public CSSStyleDeclarationImpl getDeclaration() {
        return this.declaration;
    }

    public boolean isBlock() {
        return this.block;
    }
}

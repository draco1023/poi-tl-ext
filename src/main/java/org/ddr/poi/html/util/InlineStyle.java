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

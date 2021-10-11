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
 * 记录表格单元格跨行列的帮助类
 *
 * @author Draco
 * @since 2021-01-26
 */
public class Span {
    private int row;
    private int column;
    private boolean enabled;
    private CSSStyleDeclarationImpl style;

    public Span(int row, int column, boolean enabled, CSSStyleDeclarationImpl style) {
        this.row = row;
        this.column = column;
        this.enabled = enabled;
        this.style = style;
    }

    public int getRow() {
        return this.row;
    }

    public int getColumn() {
        return this.column;
    }

    public boolean isEnabled() {
        return this.enabled;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public void setColumn(int column) {
        this.column = column;
    }

    public void setEnabled(boolean enabled) {
        this.enabled = enabled;
    }

    public CSSStyleDeclarationImpl getStyle() {
        return style;
    }

    public String toString() {
        return "Span(row=" + this.getRow() + ", column=" + this.getColumn() + ", enabled=" + this.isEnabled() + ")";
    }
}

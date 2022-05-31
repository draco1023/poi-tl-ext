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

import org.ddr.poi.html.util.CSSLength;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLevelSuffix;

import java.util.List;

/**
 * @author Draco
 * @since 2021-10-26
 */
public class HtmlRenderConfig {
    private String globalFont;

    private CSSLength globalFontSize;
    private int globalFontSizeInHalfPoints;

    private boolean showDefaultTableBorderInTableCell;
    private List<ElementRenderer> customRenderers;

    private int numberingIndent = -1;
    private STLevelSuffix.Enum numberingSpacing;

    /**
     * @return global font family
     */
    public String getGlobalFont() {
        return globalFont;
    }

    public void setGlobalFont(String globalFont) {
        this.globalFont = globalFont;
    }

    /**
     * @return global font size
     */
    public CSSLength getGlobalFontSize() {
        return globalFontSize;
    }

    public void setGlobalFontSize(CSSLength globalFontSize) {
        this.globalFontSize = globalFontSize;
        globalFontSizeInHalfPoints = globalFontSize == null ? 0 : globalFontSize.toHalfPoints();
    }

    public int getGlobalFontSizeInHalfPoints() {
        return globalFontSizeInHalfPoints;
    }

    /**
     * @return whether to show default table borders if the table inside a table cell
     */
    public boolean isShowDefaultTableBorderInTableCell() {
        return showDefaultTableBorderInTableCell;
    }

    public void setShowDefaultTableBorderInTableCell(boolean showDefaultTableBorderInTableCell) {
        this.showDefaultTableBorderInTableCell = showDefaultTableBorderInTableCell;
    }

    /**
     * @return custom html tag renderers
     */
    public List<ElementRenderer> getCustomRenderers() {
        return customRenderers;
    }

    public void setCustomRenderers(List<ElementRenderer> customRenderers) {
        this.customRenderers = customRenderers;
    }

    /**
     * @return custom numbering indent
     */
    public int getNumberingIndent() {
        return numberingIndent;
    }

    public void setNumberingIndent(int numberingIndent) {
        this.numberingIndent = numberingIndent;
    }

    /**
     * @return custom numbering spacing
     */
    public STLevelSuffix.Enum getNumberingSpacing() {
        return numberingSpacing;
    }

    public void setNumberingSpacing(STLevelSuffix.Enum numberingSpacing) {
        this.numberingSpacing = numberingSpacing;
    }
}

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

import java.util.Map;
import java.util.Objects;

/**
 * colspan对应的CSS长度值
 *
 * @author Draco
 * @since 2021-10-19
 */
public class SpanWidth extends CSSLength {
    private final int span;
    private final int column;
    private final CSSLength[] lengths;
    private final boolean explicitWidth;

    public SpanWidth(CSSLength length, int column, int span, boolean explicitWidth) {
        super(length.getValue(), length.getUnit());
        this.column = column;
        this.span = span;
        lengths = new CSSLength[span];
        this.explicitWidth = explicitWidth;
    }

    public void setLength(Map<Integer, CSSLength> map) {
        if (!isValid()) {
            for (int i = 0; i < span; i++) {
                map.putIfAbsent(column + i, CSSLength.INVALID);
            }
            return;
        }
        int invalidCount = 0;
        boolean percent = isPercent();
        double total = percent ? getValue() : unitValue();
        CSSLengthUnit unit = percent ? CSSLengthUnit.PERCENT : CSSLengthUnit.EMU;
        for (int i = 0; i < span; i++) {
            Integer index = column + i;
            CSSLength length = map.getOrDefault(index, CSSLength.INVALID);
            lengths[i] = length;
            if (length.isValid()) {
                if (percent && length.isPercent()) {
                    total -= length.getValue();
                } else if (!percent && !length.isPercent()) {
                    total -= length.unitValue();
                } else {
                    lengths[i] = CSSLength.INVALID;
                    invalidCount++;
                }
            } else {
                invalidCount++;
            }
        }
        if (invalidCount > 0) {
            final CSSLength length = new CSSLength(total / invalidCount, unit);
            for (int i = 0; i < span; i++) {
                if (!lengths[i].isValid()) {
                    map.compute(column + i, (key, value) -> {
                        if (value == null || !value.isValid()) {
                            return length;
                        } else if (explicitWidth ^ percent) {
                            return length;
                        }
                        return value;
                    });
                }
            }
        }
    }

    public int getSpan() {
        return span;
    }

    public int getColumn() {
        return column;
    }

    public CSSLength[] getLengths() {
        return lengths;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        if (!super.equals(o)) return false;
        SpanWidth spanWidth = (SpanWidth) o;
        return span == spanWidth.span &&
                column == spanWidth.column &&
                super.equals(o);
    }

    @Override
    public int hashCode() {
        return Objects.hash(column, span);
    }
}

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

import java.util.Arrays;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * CSS长度值
 *
 * @author Draco
 * @since 2021-03-01
 */
public class CSSLength {
    private static final Pattern LENGTH_PATTERN = Pattern.compile(
            "(-?\\d+\\.?\\d*)(" +
                    Arrays.stream(CSSLengthUnit.values()).filter(CSSLengthUnit::isSystem)
                            .map(CSSLengthUnit::getLiteral).collect(Collectors.joining("|"))
                    + ")"
    );

    public static final CSSLength INVALID = new CSSLength(Double.NaN, null);

    private double value;
    private CSSLengthUnit unit;

    public CSSLength(double value, CSSLengthUnit unit) {
        this.value = value;
        this.unit = unit;
    }

    public double getValue() {
        return value;
    }

    public CSSLengthUnit getUnit() {
        return unit;
    }

    @Override
    public String toString() {
        return String.format("%.2f", value) + unit.getLiteral();
    }

    /**
     * 单位转EMU，适用于绝对长度单位
     */
    public int toEMU() {
        validate();
        requireAbsoluteUnit();
        return (int) Math.rint(unitValue());
    }

    public double unitValue() {
        return value * unit.absoluteFactor();
    }

    private void requireAbsoluteUnit() {
        if (unit.isRelative()) {
            throw new UnsupportedOperationException("Can not convert a relative length to EMU: " + toString());
        }
    }

    private void validate() {
        if (!isValid()) {
            throw new UnsupportedOperationException("Invalid CSS length");
        }
    }

    public CSSLength to(CSSLengthUnit other) {
        validate();
        return new CSSLength(value * unit.to(other), other);
    }

    public int toHalfPoints() {
        validate();
        return (int) Math.rint(value * unit.to(CSSLengthUnit.PT) * 2);
    }

    public boolean isValid() {
        return unit != null && !Double.isNaN(value) && !Double.isInfinite(value);
    }

    public boolean isPercent() {
        return unit == CSSLengthUnit.PERCENT;
    }

    public boolean isValidPercent() {
        return isValid() && isPercent();
    }

    public static CSSLength of(String text) {
        if (text == null || text.isEmpty()) {
            return INVALID;
        }
        Matcher matcher = LENGTH_PATTERN.matcher(text);
        if (matcher.matches()) {
            return new CSSLength(Double.parseDouble(matcher.group(1)), CSSLengthUnit.of(matcher.group(2)));
        }
        return INVALID;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        CSSLength other = (CSSLength) o;
        if (!isValid() && !other.isValid()) return true;
        return Double.compare(other.value, value) == 0 &&
                unit == other.unit;
    }

    @Override
    public int hashCode() {
        return Objects.hash(value, unit);
    }
}

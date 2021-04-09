package org.ddr.poi.html.util;

import org.apache.poi.util.Units;
import org.ddr.poi.html.HtmlConstants;

import java.util.Arrays;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 长度单位
 *
 * @author Draco
 * @since 2021-03-01
 */
public enum CSSLengthUnit {
    CM(HtmlConstants.CM, true, false, false, Units.EMU_PER_CENTIMETER, 1, 1),
    MM(HtmlConstants.MM, true, false, false, Units.EMU_PER_CENTIMETER, 10, -1),
    IN(HtmlConstants.IN, true, false, false, Units.EMU_PER_POINT, 72, 1),
    PX(HtmlConstants.PX, true, false, false, Units.EMU_PER_PIXEL, 1, 1),
    PT(HtmlConstants.PT, true, false, false, Units.EMU_PER_POINT, 1, 1),
    PC(HtmlConstants.PC, true, false, false, Units.EMU_PER_POINT, 12, 1),

    EMU(HtmlConstants.EMU, false, false, false, 1, 1, 1),
    TWIP(HtmlConstants.TWIP, false, false, false, Units.EMU_PER_POINT, 20, -1),

    REM(HtmlConstants.REM, true, true, false, 1, 1, 1),
    EM(HtmlConstants.EM, true, true, true, 1, 1, 1),
    VW(HtmlConstants.VW, true, true, false, 1, 100, -1),
    VH(HtmlConstants.VH, true, true, false, 1, 100, -1),
    VMIN(HtmlConstants.VMIN, true, true, false, 1, 100, -1),
    VMAX(HtmlConstants.VMAX, true, true, false, 1, 100, -1),

    PERCENT(HtmlConstants.PERCENT, true, true, true, 1, 100, -1);

    private static final Map<String, CSSLengthUnit> LITERAL_MAP = Arrays.stream(values())
            .filter(CSSLengthUnit::isSystem)
            .collect(Collectors.toMap(CSSLengthUnit::getLiteral, Function.identity()));

    /**
     * 单位字面值
     */
    private final String literal;
    /**
     * 是否为系统单位，该枚举中包含部分自定义单位以便换算
     */
    private final boolean system;
    /**
     * 是否为相对长度
     */
    private final boolean relative;
    /**
     * 是否相对父元素，false表示相对于根元素
     */
    private final boolean relativeToParent;
    // 下面3个属性联合表示单位系数，绝对长度以EMU为基准，相对长度以1为基准
    private final int unit;
    private final int factor;
    private final int power;

    CSSLengthUnit(String literal, boolean system, boolean relative, boolean relativeToParent, int unit, int factor, int power) {
        this.literal = literal;
        this.system = system;
        this.relative = relative;
        this.relativeToParent = relativeToParent;
        this.unit = unit;
        this.factor = factor;
        this.power = power;
    }

    public String getLiteral() {
        return literal;
    }

    public boolean isSystem() {
        return system;
    }

    public boolean isRelative() {
        return relative;
    }

    public boolean isRelativeToParent() {
        return relativeToParent;
    }

    @Override
    public String toString() {
        return literal;
    }

    public double absoluteFactor() {
        return unit * Math.pow(factor, power);
    }

    public double to(CSSLengthUnit other) {
        Objects.requireNonNull(other, "Target CSS length unit must not be null");
        if (this.isRelative()) {
            throw new IllegalArgumentException("Can not convert from a relative unit");
        }
        if (other.isRelative()) {
            throw new IllegalArgumentException("Can not convert to a relative unit");
        }
        return absoluteFactor() / other.absoluteFactor();
    }

    /**
     * 单位字面值转为单位
     *
     * @param literal 单位字面值
     * @return 长度单位
     */
    public static CSSLengthUnit of(String literal) {
        return LITERAL_MAP.get(literal);
    }
}

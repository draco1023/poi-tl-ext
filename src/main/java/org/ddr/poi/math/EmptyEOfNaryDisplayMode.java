package org.ddr.poi.math;

/**
 * 空的N元组显示模式
 *
 * @author Draco
 * @since 2025-06-09 16:44
 */
public enum EmptyEOfNaryDisplayMode {
    /**
     * 默认：显示输入框
     */
    DEFAULT(0x00),
    /**
     * 零宽度
     */
    ZERO_WIDTH(0x01),
    /**
     * 隐藏输入框
     */
    HIDDEN(0x10),
    /**
     * 零宽度并且隐藏
     */
    ZERO_WIDTH_HIDDEN(0x11);

    private final int value;

    EmptyEOfNaryDisplayMode(int value) {
        this.value = value;
    }

    public int getValue() {
        return value;
    }

    public boolean isZeroWidth() {
        return (value & ZERO_WIDTH.getValue()) != 0;
    }

    public boolean isHidden() {
        return (value & HIDDEN.getValue()) != 0;
    }
}

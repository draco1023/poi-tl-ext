package org.ddr.poi.html.util;

/**
 * 列表样式
 *
 * @author Draco
 * @since 2024-09-12
 */
public class ListStyle {
    private final ListStyleType numberFormat;

    private final boolean hanging;

    private final int left;

    private final int right;

    public ListStyle(ListStyleType numberFormat, boolean hanging, int left, int right) {
        this.numberFormat = numberFormat;
        this.hanging = hanging;
        this.left = left;
        this.right = right;
    }

    public ListStyleType getNumberFormat() {
        return numberFormat;
    }

    public boolean isHanging() {
        return hanging;
    }

    public int getLeft() {
        return left;
    }

    public int getRight() {
        return right;
    }

    @Override
    public String toString() {
        return "ListStyle{" +
                "numberFormat=" + numberFormat +
                ", hanging=" + hanging +
                ", left=" + left +
                ", right=" + right +
                '}';
    }
}

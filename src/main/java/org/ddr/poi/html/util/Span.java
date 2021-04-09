package org.ddr.poi.html.util;

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

    public Span(int row, int column, boolean enabled) {
        this.row = row;
        this.column = column;
        this.enabled = enabled;
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

    public String toString() {
        return "Span(row=" + this.getRow() + ", column=" + this.getColumn() + ", enabled=" + this.isEnabled() + ")";
    }
}

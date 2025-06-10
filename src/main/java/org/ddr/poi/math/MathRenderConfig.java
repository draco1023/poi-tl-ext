package org.ddr.poi.math;

/**
 * 公式渲染配置
 * @author Draco
 * @since 2025-06-09 16:08
 */
public class MathRenderConfig {
    private EmptyEOfNaryDisplayMode emptyEOfNaryDisplayMode = EmptyEOfNaryDisplayMode.DEFAULT;

    public EmptyEOfNaryDisplayMode getEmptyEOfNaryOption() {
        return emptyEOfNaryDisplayMode;
    }

    public void setEmptyEOfNaryOption(EmptyEOfNaryDisplayMode emptyEOfNaryDisplayMode) {
        this.emptyEOfNaryDisplayMode = emptyEOfNaryDisplayMode;
    }
}

package org.ddr.poi.html.util;

import org.apache.poi.xwpf.usermodel.IRunElement;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.lang.reflect.Field;
import java.util.List;

/**
 * Wrapper class for XWPFParagraph runs
 *
 * @author Draco
 * @since 2022-07-05
 */
public class XWPFParagraphRuns {
    private static Field runsField;
    private static Field irunsField;

    static {
        try {
            runsField = XWPFParagraph.class.getDeclaredField("runs");
            irunsField = XWPFParagraph.class.getDeclaredField("iruns");
        } catch (NoSuchFieldException ignored) {
        }
        runsField.setAccessible(true);
        irunsField.setAccessible(true);
    }

    private List<XWPFRun> runs;
    private List<IRunElement> iruns;

    @SuppressWarnings("unchecked")
    public XWPFParagraphRuns(XWPFParagraph paragraph) {
        try {
            runs = (List<XWPFRun>) runsField.get(paragraph);
            iruns = (List<IRunElement>) irunsField.get(paragraph);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Remove run at position without modifying xml
     *
     * @param pos index of run
     */
    public void remove(int pos) {
        XWPFRun run = runs.remove(pos);
        iruns.remove(run);
    }

    /**
     * @return runs count
     */
    public int runCount() {
        return runs.size();
    }
}

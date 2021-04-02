package com.mng.application;

import com.mng.domain.TextFormat;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.beans.factory.annotation.Autowired;

import java.util.List;

public class TextFormatApplication {
    @Autowired
    TextFormat textFormat = new TextFormat();

    public boolean isCorrectSize(XWPFDocument document, XWPFParagraph paragraph, Float size) {
        TextFormat t = new TextFormat();
        List<XWPFRun> runs = paragraph.getRuns();
        for (XWPFRun r : runs) {
            if (!r.getText(0).equals(" ")) {
                try {
                    if (textFormat.getFontSize(document, paragraph, r) != size) {
                        return false;
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        return true;
    }

    public boolean isCorrectTheme(XWPFDocument document, XWPFParagraph paragraph, String theme) {
        List<XWPFRun> runs = paragraph.getRuns();
        for (XWPFRun r : runs) {
            if (!r.getText(0).equals(" ")) {
                try {
                    if (!textFormat.getFontTheme(r, document, paragraph).equals(theme) ||
                            textFormat.getFontTheme(r, document, paragraph) == null) {
                        return false;
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        return true;
    }
}

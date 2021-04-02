package com.mng.domain;

import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.regex.Pattern;

public class TextType {
    /**
     * 对空格、符号、数字不区分中英文
     *
     * @param str
     * @return
     */
    public String removeExtraString(String str) {
        return str.replaceAll(" ", "")
                .replaceAll("[\\pP\\p{Punct}]", "")
                .replaceAll("[0-9]*", "");

    }

    /**
     * 判断是否英文
     *
     * @param xwpfRun
     * @return
     */
    public Boolean isEnglishFont(XWPFRun xwpfRun) {
        if (this.removeExtraString(xwpfRun.getText(0)).matches("[a-zA-Z]+")) {
            return true;
        }
        return false;
    }

    /**
     * 判断是否中文
     *
     * @param xwpfRun
     * @return
     */
    public Boolean isChinessFont(XWPFRun xwpfRun) {
        Pattern pattern = Pattern.compile("[\\u4e00-\\u9fa5]+");
        if (pattern.matcher(this.removeExtraString(xwpfRun.getText(0))).find()) {
            return true;
        }
        return false;
    }
}

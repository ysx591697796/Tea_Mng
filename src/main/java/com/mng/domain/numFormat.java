package com.mng.domain;

import com.mng.application.TextFormatApplication;

public class numFormat {
    public void isNum(){
        System.out.println(TextFormatApplication.wordFontSize.get("asd"));
        String reg = "^[0-9]+(.[0-9]+)?$";
        String str = "sad";
        System.out.println(str.matches(reg));
    }
}

package com.mng.application;

import com.mng.domain.ParasFormat;
import com.mng.domain.TextFormat;
import com.mng.domain.TextType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.beans.factory.annotation.Autowired;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class TextFormatApplication {

    @Autowired
    TextFormat textFormat = new TextFormat();

    @Autowired
    ParasFormat pf = new ParasFormat();

    public final static Map<String, Float> wordFontSize = new HashMap<String, Float>() {{
        put("初号", (float) 42.0);
        put("小初", (float) 36.0);
        put("一号", (float) 26.0);
        put("小一", (float) 24.0);
        put("二号", (float) 22.0);
        put("小二", (float) 18.0);
        put("三号", (float) 16.0);
        put("小三", (float) 15.0);
        put("四号", (float) 14.0);
        put("小四", (float) 12.0);
        put("五号", (float) 10.5);
        put("小五", (float) 9.0);
        put("六号", (float) 7.5);
        put("小六", (float) 6.5);
        put("七号", (float) 5.5);
        put("八号", (float) 5.0);
    }};

    public static String chnHeaderTheme = "黑体";
    public static float chnHeaderSize = (float) 22.0;
    public static String chnNameTheme = "仿宋_GB2312";
    public static float chnNameSize = (float) 12.0;
    public static String chnMajorTheme = "仿宋_GB2312";
    public static float chnMajorSize = (float) 12.0;

    public static String engHeaderTheme = "Times New Roman";
    public static float engHeaderSize = (float) 22.0;
    public static String engNameTheme = "Times New Roman";
    public static float engNameSize = (float) 12.0;
    public static String engMajorTheme = "Times New Roman";
    public static float engMajorSize = (float) 10.5;

    public static String chnAbstractHeaderTheme = "黑体";
    public static String chnAbstractTextTheme = "仿宋_GB2312";
    public static float chnAbstractSize = (float) 12.0;

    public static String engAbstractHeaderTheme = "Times New Roman";
    public static String engAbstractTextTheme = "Times New Roman";
    public static float engAbstractSize = (float) 10.5;

    public static String chnKeywordsHeaderTheme = "黑体";
    public static String chnKeywordsTextTheme = "仿宋_GB2312";
    public static float chnKeywordsSize = (float) 12.0;

    public static String engKeywordsHeaderTheme = "Times New Roman";
    public static String engKeywordsTextTheme = "Times New Roman";
    public static float engKeywordsSize = (float) 10.5;

    public static String cataTheme = "黑体";
    public static float cataSize = (float) 16.0;

    public static String cataTextTheme = "宋体";
    public static float cataTextSize = (float) 12.0;

    public static String litTheme = "黑体";
    public static float litSize = (float) 10.5;

    public static String chnlitTextTheme = "宋体";
    public static String englitTextTheme = "Times New Roman";
    public static float litTextSize = (float) 10.5;

    public static String textHeader1Theme = "黑体";
    public static float textHeader1Size = (float) 12.0;

    public static String textHeader2Theme = "宋体";
    public static float textHeader2Size = (float) 12.0;

    public static String textTheme = "宋体";
    public static float textSize = (float) 12.0;

    private XWPFDocument document;
    private List<XWPFParagraph> paragraphs;

    public TextFormatApplication() {

    }

    public TextFormatApplication(XWPFDocument document) {
        this.document = document;
        paragraphs = document.getParagraphs();
    }

    public boolean isCorrectSize(XWPFDocument document, XWPFParagraph paragraph, Float size) {
        TextFormat t = new TextFormat();
        List<XWPFRun> runs = paragraph.getRuns();
        for (XWPFRun r : runs) {
            if (r.getText(0) != null) {
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
        }
        return true;
    }

    public boolean isCorrectTheme(XWPFDocument document, XWPFParagraph paragraph, String theme) {
        List<XWPFRun> runs = paragraph.getRuns();
        for (XWPFRun r : runs) {
            if (r.getText(0) != null) {
                if (!r.getText(0).trim().equals("")) {
                    try {
                        String s = textFormat.getFontTheme(r, document, paragraph);
                        if (textFormat.getFontTheme(r, document, paragraph) == null ||
                                !textFormat.getFontTheme(r, document, paragraph).equals(theme)
                        ) {
                            return false;
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            }
        }
        return true;
    }

    public boolean isCenter(XWPFDocument document, XWPFParagraph paragraph) {

        if (paragraph.getCTP().getPPr() != null && paragraph.getCTP().getPPr().getJc() != null) {
            if (paragraph.getCTP().getPPr().getJc().getVal().toString().equals("center"))
                return true;
        }

        if (document.getStyles().getStyle(paragraph.getStyleID()) != null) {
            if (document.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr() != null &&
                    document.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getJc() != null) {
                if (document.getStyles().getStyle(paragraph.getStyleID()).getCTStyle()
                        .getPPr().getJc().getVal().toString().equals("center"))
                    return true;
            }
            if (document.getStyles().getStyle(paragraph.getStyleID()).getLinkStyleID() != null) {
                if (document.getStyles().getStyle(document.getStyles().getStyle(paragraph.getStyleID())
                        .getLinkStyleID()).getCTStyle().getPPr() != null &&
                        document.getStyles().getStyle(document.getStyles().getStyle(paragraph.getStyleID())
                                .getLinkStyleID()).getCTStyle().getPPr().getJc() != null) {
                    if (document.getStyles().getStyle(document.getStyles().getStyle(paragraph.getStyleID())
                            .getLinkStyleID()).getCTStyle()
                            .getPPr().getJc().getVal().toString().equals("center"))
                        return true;
                }
            }
        }

        return false;
    }

    public boolean isLeft(XWPFDocument document, XWPFParagraph paragraph) {

        if (paragraph.getCTP().getPPr() != null && paragraph.getCTP().getPPr().getJc() != null) {
            if (paragraph.getCTP().getPPr().getJc().getVal().toString().equals("center") ||
                    paragraph.getCTP().getPPr().getJc().getVal().toString().equals("right"))
                return false;
        }

        if (document.getStyles().getStyle(paragraph.getStyleID()) != null) {
            if (document.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr() != null &&
                    document.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getJc() != null) {
                if (document.getStyles().getStyle(paragraph.getStyleID()).getCTStyle()
                        .getPPr().getJc().getVal().toString().equals("center") ||
                        document.getStyles().getStyle(paragraph.getStyleID()).getCTStyle()
                                .getPPr().getJc().getVal().toString().equals("right"))
                    return false;
            }
            if (document.getStyles().getStyle(paragraph.getStyleID()).getLinkStyleID() != null) {
                if (document.getStyles().getStyle(document.getStyles().getStyle(paragraph.getStyleID())
                        .getLinkStyleID()).getCTStyle().getPPr() != null &&
                        document.getStyles().getStyle(document.getStyles().getStyle(paragraph.getStyleID())
                                .getLinkStyleID()).getCTStyle().getPPr().getJc() != null) {
                    if (document.getStyles().getStyle(document.getStyles().getStyle(paragraph.getStyleID())
                            .getLinkStyleID()).getCTStyle()
                            .getPPr().getJc().getVal().toString().equals("center") || document.getStyles().getStyle(document.getStyles().getStyle(paragraph.getStyleID())
                            .getLinkStyleID()).getCTStyle()
                            .getPPr().getJc().getVal().toString().equals("right"))
                        return false;
                }
            }
        }

        return true;
    }

    //0对,1摘要字体错,2内容字体错
    public int isCorrectAbstractTheme1(XWPFDocument document, XWPFParagraph paragraph, String theme1, String theme2) {
        boolean aFlag = true;
        boolean tFlag = true;
        List<XWPFRun> runs = paragraph.getRuns();
        try {
            if (textFormat.getFontTheme(runs.get(0), document, paragraph) == null ||
                    !textFormat.getFontTheme(runs.get(0), document, paragraph).equals(theme1)
            ) {
                aFlag = false;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        for (int i = 3; i < runs.size(); i++) {
            if (!runs.get(i).getText(0).equals(" ")) {
                try {
                    if (textFormat.getFontTheme(runs.get(i), document, paragraph) == null ||
                            !textFormat.getFontTheme(runs.get(i), document, paragraph).equals(theme2)
                    ) {
                        tFlag = false;
                        break;
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        if (!aFlag || !tFlag) {
            if (!aFlag && !tFlag) return 3;
            if (!aFlag) return 1;
            if (!tFlag) return 2;
        }
        return 0;
    }

    //0对,1摘要字体错,2内容字体错
    public int isCorrectAbstractTheme2(XWPFDocument document, XWPFParagraph paragraph, String theme1, String theme2) {
        boolean aFlag = true;
        boolean tFlag = true;
        List<XWPFRun> runs = paragraph.getRuns();
        try {
            if (textFormat.getFontTheme(runs.get(0), document, paragraph) == null ||
                    !textFormat.getFontTheme(runs.get(0), document, paragraph).equals(theme1)
            ) {
                aFlag = false;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        for (int i = 4; i < runs.size(); i++) {
            if (!runs.get(i).getText(0).equals(" ")) {
                try {
                    if (textFormat.getFontTheme(runs.get(i), document, paragraph) == null ||
                            !textFormat.getFontTheme(runs.get(i), document, paragraph).equals(theme2)
                    ) {
                        tFlag = false;
                        break;
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        if (!aFlag || !tFlag) {
            if (!aFlag && !tFlag) return 3;
            if (!aFlag) return 1;
            if (!tFlag) return 2;
        }
        return 0;
    }

    //页大小 A4:11906W,16838H
    public String pageSizePart() {
        StringBuilder pageStr = new StringBuilder();
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        CTPageSz pageSz = sectPr.getPgSz();
        if (!pageSz.getW().toString().equals("11906") || !pageSz.getH().toString().equals("16838")) {
            pageStr.append("-页面大小- 页面大小出错，需设置A4!\n");
        }
        return pageStr.toString();
    }

    //页边距 top:1701(30mm) bottom:1418(25mm) left:1701(30mm) right:1134(20mm)
    public String pageMarPart() {
        StringBuilder marStr = new StringBuilder();
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        CTPageMar ctPageMar = sectPr.getPgMar();
        if (!ctPageMar.getTop().toString().equals("1701") || !ctPageMar.getBottom().toString().equals("1418") ||
                !ctPageMar.getLeft().toString().equals("1701") || !ctPageMar.getRight().toString().equals("1134")) {
            marStr.append("-页面边距- 页面边距出错，");
            if (!ctPageMar.getTop().toString().equals("1701")) {
                marStr.append("上(需设置30mm)");
            }
            if (!ctPageMar.getBottom().toString().equals("1418")) {
                marStr.append("下(需设置25mm)");
            }
            if (!ctPageMar.getLeft().toString().equals("1701")) {
                marStr.append("左(需设置30mm)");
            }
            if (!ctPageMar.getRight().toString().equals("1134")) {
                marStr.append("有(需设置20mm)");
            }
            marStr.append("!\n");
        }
        return marStr.toString();
    }

    //页码
    public String footerPart() {
        StringBuilder footerStr = new StringBuilder();
        List<XWPFFooter> footers = document.getFooterList();
        if (footers.size() != 0) {
            for (XWPFFooter f : footers
            ) {
                if (f._getHdrFtr().getSdtList() != null && f._getHdrFtr().getSdtList().size() > 0) {
                    if (f._getHdrFtr().getSdtList().get(0).getSdtContent() != null) {
                        boolean centerFlag = true;
                        if (f._getHdrFtr().getSdtList().get(0).getSdtContent().getPList().get(0).getPPr() != null
                                && f._getHdrFtr().getSdtList().get(0).getSdtContent().getPList().get(0).getPPr().getJc() != null) {
                            if (!f._getHdrFtr().getSdtList().get(0).getSdtContent().getPList().get(0).getPPr().getJc().getVal().toString().equals("center")) {
                                centerFlag = false;
                            }
                        }
                        boolean themeFlag = true;
                        if (f._getHdrFtr().getSdtList().get(0).getSdtContent().getPList().get(0).getPPr() != null
                                && f._getHdrFtr().getSdtList().get(0).getSdtContent().getPList().get(0).getPPr().getPStyle() != null) {
                            String temp = f._getHdrFtr().getSdtList().get(0).getSdtContent().getPList().get(0).getPPr().getPStyle().getVal();
                            if (document.getStyles().getStyle(temp) != null) {
                                if (document.getStyles().getStyle(temp).getCTStyle().getRPr() != null) {
                                    if (document.getStyles().getStyle(temp).getCTStyle().getRPr().getRFonts() != null) {
                                        CTFonts rFonts1 = document.getStyles().getStyle(temp).getCTStyle().getRPr().getRFonts();
                                        if (rFonts1.getAscii() != null) {
                                            if (!rFonts1.getAscii().equals("Times New Roman")) {
                                                themeFlag = false;
                                            }
                                        }
                                    }
                                }
                                if (document.getStyles().getStyle(temp).getLinkStyleID() != null) {
                                    if (document.getStyles().getStyle(document.getStyles().getStyle(temp)
                                            .getLinkStyleID()).getCTStyle().getRPr() != null) {
                                        if (document.getStyles().getStyle(document.getStyles().getStyle(temp)
                                                .getLinkStyleID()).getCTStyle().getRPr().getRFonts() != null) {
                                            TextType ty = new TextType();
                                            CTFonts rFonts1 = document.getStyles().getStyle(document.getStyles().getStyle(temp)
                                                    .getLinkStyleID()).getCTStyle().getRPr().getRFonts();
                                            if (rFonts1.getAscii() != null) {
                                                if (!rFonts1.getAscii().equals("Times New Roman")) {
                                                    themeFlag = false;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        boolean sizeFlag = true;
                        if (f._getHdrFtr().getSdtList().get(0).getSdtContent().getPList().get(0).getPPr() != null
                                && f._getHdrFtr().getSdtList().get(0).getSdtContent().getPList().get(0).getPPr().getPStyle() != null) {
                            String temp = f._getHdrFtr().getSdtList().get(0).getSdtContent().getPList().get(0).getPPr().getPStyle().getVal();
                            if (document.getStyles().getStyle(temp) != null) {
                                if (document.getStyles().getStyle(temp).getCTStyle().getRPr() != null) {
                                    CTRPr rPr = document.getStyles().getStyle(temp).getCTStyle().getRPr();
                                    if (rPr.getSz() != null) {
                                        if ((float) rPr.getSz().getVal().longValue() / 2 != 9.0)
                                            sizeFlag = false;
                                    }
                                }
                                if (document.getStyles().getStyle(temp).getLinkStyleID() != null) {
                                    if (document.getStyles().getStyle(document.getStyles().getStyle(temp)
                                            .getLinkStyleID()).getCTStyle().getRPr() != null) {
                                        CTRPr rPr = document.getStyles().getStyle(document.getStyles().getStyle(temp)
                                                .getLinkStyleID()).getCTStyle().getRPr();
                                        if (rPr.getSz() != null) {
                                            if ((float) rPr.getSz().getVal().longValue() / 2 != 9.0)
                                                sizeFlag = false;
                                        }
                                    }
                                }
                            }
                        }
                        if (!themeFlag || !sizeFlag || !centerFlag) {
                            if (!sizeFlag && !themeFlag && !centerFlag) {
                                footerStr.append("-页码- 段落字体大小及样式出错且未居中!\n");
                            } else if (!themeFlag && !centerFlag) {
                                footerStr.append("-页码- 段落字体样式出错且未居中!\n");
                            } else if (!sizeFlag && !centerFlag) {
                                footerStr.append("-页码- 段落字体大小出错且未居中!\n");
                            } else if (!sizeFlag && !themeFlag) {
                                footerStr.append("-页码- 段落字体大小及样式出错!\n");
                            } else if (!themeFlag) {
                                footerStr.append("-页码- 段落字体样式出错!\n");
                            } else if (!sizeFlag) {
                                footerStr.append("-页码- 段落字体大小出错!\n");
                            } else if (!centerFlag) {
                                footerStr.append("-页码- 未居中!\n");
                            }
                        }
                    } else {
                        footerStr.append("无法检测到页码!\n");
                    }
                } else if (f.getText() != null && !f.getText().trim().equals(" ") && f.getText().length() > 1) {
                    List<XWPFParagraph> xwpfParagraphs = f.getParagraphs();
                    if (xwpfParagraphs.size() != 0) {
                        for (XWPFParagraph p : xwpfParagraphs
                        ) {
                            boolean centerFlag = this.isCenter(document, p);
                            boolean themeFlag = this.
                                    isCorrectTheme(document, p, "Times New Roman");
                            boolean sizeFlag = this.isCorrectSize(document, p, (float) 9.0);
                            if (!themeFlag || !sizeFlag || !centerFlag) {
                                if (!sizeFlag && !themeFlag && !centerFlag) {
                                    footerStr.append("-页码- [" + p.getText().trim()
                                            + "] 段落字体大小及样式出错且未居中!\n");
                                } else if (!themeFlag && !centerFlag) {
                                    footerStr.append("-页码- [" + p.getText().trim()
                                            + "] 段落字体样式出错且未居中!\n");
                                } else if (!sizeFlag && !centerFlag) {
                                    footerStr.append("-页码- [" + p.getText().trim()
                                            + "] 段落字体大小出错且未居中!\n");
                                } else if (!sizeFlag && !themeFlag) {
                                    footerStr.append("-页码- [" + p.getText().trim()
                                            + "] 段落字体大小及样式出错!\n");
                                } else if (!themeFlag) {
                                    footerStr.append("-页码- [" + p.getText().trim()
                                            + "] 段落字体样式出错!\n");
                                } else if (!sizeFlag) {
                                    footerStr.append("-页码- [" + p.getText().trim()
                                            + "] 段落字体大小出错!\n");
                                } else if (!centerFlag) {
                                    footerStr.append("-页码- [" + p.getText().trim()
                                            + "] 段落未居中!\n");
                                }
                            }
                        }
                    } else {
                        footerStr.append("无法检测到页码!\n");
                    }
                }
            }
        }
        return footerStr.toString();
    }

    //中文标题部分
    public StringBuilder chnHeaderPart(int i) {
        StringBuilder paperHStr = new StringBuilder();
        //专业段
        if (paragraphs.get(i).getText().startsWith("专业")) {
            //论文标题部分
            if (!paragraphs.get(i - 3).getText().trim().equals("")) {
                boolean centerFlag = this.isCenter(document, paragraphs.get(i - 3));
                boolean themeFlag = this.
                        isCorrectTheme(document, paragraphs.get(i - 3), chnHeaderTheme);
                //二号
                boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i - 3)
                        , chnHeaderSize);
                if (!themeFlag || !sizeFlag || !centerFlag) {
                    paperHStr.append("-论文标题(中文)- ");
                    if (!themeFlag) {
                        paperHStr.append("字体样式出错、");
                    }
                    if (!sizeFlag) {
                        paperHStr.append("字体大小出错、");
                    }
                    if (!centerFlag) {
                        paperHStr.append("未居中");
                    }
                    if (paperHStr.toString().endsWith("、")) {
                        paperHStr.deleteCharAt(paperHStr.length() - 1);
                        paperHStr.append("!\n");
                    } else {
                        paperHStr.append("!\n");
                    }
                }
                if (!paragraphs.get(i - 2).getText().trim().equals("")) {
                    paperHStr.append("-论文标题(中文)- 下面应空一行!\n");
                }
            } else {
                paperHStr.append("-论文标题(中文)- 标题位置出错!\n");
            }

            //姓名部分
            if (!paragraphs.get(i - 1).getText().trim().equals("")) {
                boolean centerFlag = this.isCenter(document, paragraphs.get(i - 1));
                boolean themeFlag = this.
                        isCorrectTheme(document, paragraphs.get(i - 1), chnNameTheme);
                boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i - 1)
                        , chnNameSize);
                if (!themeFlag || !sizeFlag || !centerFlag) {
                    paperHStr.append("-姓名(中文)- ");
                    if (!themeFlag) {
                        paperHStr.append("字体样式出错、");
                    }
                    if (!sizeFlag) {
                        paperHStr.append("字体大小出错、");
                    }
                    if (!centerFlag) {
                        paperHStr.append("未居中");
                    }
                    if (paperHStr.toString().endsWith("、")) {
                        paperHStr.deleteCharAt(paperHStr.length() - 1);
                        paperHStr.append("!\n");
                    } else {
                        paperHStr.append("!\n");
                    }
                }
            } else {
                paperHStr.append("-姓名(中文)- 姓名位置出错!\n");
            }

            //专业行
            boolean centerFlag = this.isCenter(document, paragraphs.get(i));
            boolean themeFlag = this.
                    isCorrectTheme(document, paragraphs.get(i), chnMajorTheme);
            boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i)
                    , chnMajorSize);
            if (!themeFlag || !sizeFlag || !centerFlag) {
                paperHStr.append("-专业- ");
                if (!themeFlag) {
                    paperHStr.append("字体样式出错、");
                }
                if (!sizeFlag) {
                    paperHStr.append("字体大小出错、");
                }
                if (!centerFlag) {
                    paperHStr.append("未居中");
                }
                if (paperHStr.toString().endsWith("、")) {
                    paperHStr.deleteCharAt(paperHStr.length() - 1);
                    paperHStr.append("!\n");
                } else {
                    paperHStr.append("!\n");
                }
            }

            //词前空格
            try {
                int index = paragraphs.get(i).getText().indexOf("学号");
                if (!Character.isWhitespace(paragraphs.get(i).getText().charAt(index - 1))
                        || !Character.isWhitespace(paragraphs.get(i).getText().charAt(index - 2))) {
                    paperHStr.append("-学号- 词前未空两格!\n");
                }
                index = paragraphs.get(i).getText().indexOf("指导老师");
                if (!Character.isWhitespace(paragraphs.get(i).getText().charAt(index - 1))
                        || !Character.isWhitespace(paragraphs.get(i).getText().charAt(index - 2))) {
                    paperHStr.append("-指导老师- 词前未空两格!\n");
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return paperHStr;
    }

    //英文标题部分
    public StringBuilder engHeaderPart(int i) {
        StringBuilder paperHStr = new StringBuilder();
        //专业段
        if (paragraphs.get(i).getText().startsWith("Major")) {
            //论文标题部分
            if (!paragraphs.get(i - 3).getText().trim().equals("")) {
                boolean centerFlag = this.isCenter(document, paragraphs.get(i - 3));
                boolean themeFlag = this.
                        isCorrectTheme(document, paragraphs.get(i - 3), engHeaderTheme);
                //二号
                boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i - 3)
                        , engHeaderSize);
                if (!themeFlag || !sizeFlag || !centerFlag) {
                    paperHStr.append("-论文标题(英文)- ");
                    if (!themeFlag) {
                        paperHStr.append("字体样式出错、");
                    }
                    if (!sizeFlag) {
                        paperHStr.append("字体大小出错、");
                    }
                    if (!centerFlag) {
                        paperHStr.append("未居中");
                    }
                    if (paperHStr.toString().endsWith("、")) {
                        paperHStr.deleteCharAt(paperHStr.length() - 1);
                        paperHStr.append("!\n");
                    } else {
                        paperHStr.append("!\n");
                    }
                }
                if (!paragraphs.get(i - 2).getText().trim().equals("")) {
                    paperHStr.append("-论文标题(英文)- 下面应空一行!\n");
                }
            } else {
                paperHStr.append("-论文标题(英文)- 标题位置出错!\n");
            }

            //姓名部分
            if (!paragraphs.get(i - 1).getText().trim().equals("")) {
                boolean centerFlag = this.isCenter(document, paragraphs.get(i - 1));
                boolean themeFlag = this.
                        isCorrectTheme(document, paragraphs.get(i - 1), engNameTheme);
                boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i - 1)
                        , engNameSize);
                if (!themeFlag || !sizeFlag || !centerFlag) {
                    paperHStr.append("-姓名(英文)- ");
                    if (!themeFlag) {
                        paperHStr.append("字体样式出错、");
                    }
                    if (!sizeFlag) {
                        paperHStr.append("字体大小出错、");
                    }
                    if (!centerFlag) {
                        paperHStr.append("未居中");
                    }
                    if (paperHStr.toString().endsWith("、")) {
                        paperHStr.deleteCharAt(paperHStr.length() - 1);
                        paperHStr.append("!\n");
                    } else {
                        paperHStr.append("!\n");
                    }
                }
            } else {
                paperHStr.append("-姓名(英文)- 姓名位置出错!\n");
            }

            //专业行
            boolean centerFlag = this.isCenter(document, paragraphs.get(i));
            XWPFParagraph P = paragraphs.get(i);
            boolean themeFlag = this.
                    isCorrectTheme(document, paragraphs.get(i), engMajorTheme);
            boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i)
                    , engMajorSize);
            if (!themeFlag || !sizeFlag || !centerFlag) {
                paperHStr.append("-Major- ");
                if (!themeFlag) {
                    paperHStr.append("字体样式出错、");
                }
                if (!sizeFlag) {
                    paperHStr.append("字体大小出错、");
                }
                if (!centerFlag) {
                    paperHStr.append("未居中");
                }
                if (paperHStr.toString().endsWith("、")) {
                    paperHStr.deleteCharAt(paperHStr.length() - 1);
                    paperHStr.append("!\n");
                } else {
                    paperHStr.append("!\n");
                }
            }

            //词前空格
            try {
                int index = paragraphs.get(i).getText().indexOf("No");
                if (!Character.isWhitespace(paragraphs.get(i).getText().charAt(index - 1))
                        || !Character.isWhitespace(paragraphs.get(i).getText().charAt(index - 2))) {
                    paperHStr.append("-No- 词前未空两格!\n");
                }
                index = paragraphs.get(i).getText().indexOf("Tutor");
                if (!Character.isWhitespace(paragraphs.get(i).getText().charAt(index - 1))
                        || !Character.isWhitespace(paragraphs.get(i).getText().charAt(index - 2))) {
                    paperHStr.append("-Tutor- 词前未空两格!\n");
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return paperHStr;
    }

    //中文摘要部分
    public StringBuilder chnAbstractpart(int i) {
        StringBuilder abstractStr = new StringBuilder();
        boolean flag = true;
        if (!paragraphs.get(i - 1).getText().trim().equals("")) {
            flag = false;
            abstractStr.append("-中文摘要- 摘要前未空行、");
        }
        try {
            if (pf.getParaFirstLineChars(document, paragraphs.get(i)) != 0) {
                if (flag) {
                    flag = false;
                    abstractStr.append("-中文摘要- 摘要未顶格、");
                } else {
                    abstractStr.append("摘要未顶格、");
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        if (paragraphs.get(i).getText().startsWith("摘要")) {
            int themeFlag = this.isCorrectAbstractTheme1(document
                    , paragraphs.get(i), chnAbstractHeaderTheme
                    , chnAbstractTextTheme);
            if (themeFlag == 1) {
                if (flag) {
                    flag = false;
                    abstractStr.append("-中文摘要- 摘要字体错、");
                } else {
                    abstractStr.append("摘要字体错、");
                }
            } else if (themeFlag == 2) {
                if (flag) {
                    flag = false;
                    abstractStr.append("-中文摘要- 摘要内容字体错、");
                } else {
                    abstractStr.append("摘要内容字体错、");
                }
            } else if (themeFlag == 3) {
                if (flag) {
                    flag = false;
                    abstractStr.append("-中文摘要- 摘要和摘要内容字体全错、");
                } else {
                    abstractStr.append("摘要和摘要内容字体全错、");
                }
            }
        } else if (paragraphs.get(i).getText().startsWith("摘 要")) {
            int themeFlag = this.isCorrectAbstractTheme2(document
                    , paragraphs.get(i), chnAbstractHeaderTheme
                    , chnAbstractTextTheme);
            if (themeFlag == 1) {
                if (flag) {
                    flag = false;
                    abstractStr.append("-中文摘要- 摘要字体错、");
                } else {
                    abstractStr.append("摘要字体错、");
                }
            } else if (themeFlag == 2) {
                if (flag) {
                    flag = false;
                    abstractStr.append("-中文摘要- 摘要内容字体错、");
                } else {
                    abstractStr.append("摘要内容字体错、");
                }
            } else if (themeFlag == 3) {
                if (flag) {
                    flag = false;
                    abstractStr.append("-中文摘要- 摘要和摘要内容字体全错、");
                } else {
                    abstractStr.append("摘要和摘要内容字体全错、");
                }
            }
        }
        boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i)
                , chnAbstractSize);
        if (!sizeFlag) {
            if (flag) {
                flag = false;
                abstractStr.append("-中文摘要- 字体大小出错、");
            } else {
                abstractStr.append("字体大小出错、");
            }
        }
        if (flag == false) {
            abstractStr.deleteCharAt(abstractStr.length() - 1);
            abstractStr.append("!\n");
        }
        return abstractStr;
    }

    //英文摘要部分
    public StringBuilder engAbstractpart(int i) {
        StringBuilder abstractStr = new StringBuilder();
        boolean flag = true;
        if (!paragraphs.get(i - 1).getText().trim().equals("")) {
            flag = false;
            abstractStr.append("-Abstract- 摘要前未空行、");
        }
        try {
            if (pf.getParaFirstLineChars(document, paragraphs.get(i)) != 0) {
                if (flag) {
                    flag = false;
                    abstractStr.append("-Abstract- Abstract未顶格、");
                } else {
                    abstractStr.append("Abstract未顶格、");
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        try {
            if (!textFormat.getFontBold(document, paragraphs.get(i), paragraphs.get(i).getRuns().get(0))) {
                if (flag) {
                    flag = false;
                    abstractStr.append("-Abstract- Abstract未加粗、");
                } else {
                    abstractStr.append("Abstract未加粗、");
                }
            }
            if (paragraphs.get(i).getText().startsWith("Abstract")) {
                int themeFlag = this.isCorrectAbstractTheme1(document
                        , paragraphs.get(i), engAbstractHeaderTheme, engAbstractTextTheme);
                if (themeFlag == 1) {
                    if (flag) {
                        flag = false;
                        abstractStr.append("-Abstract- Abstract字体错、");
                    } else {
                        abstractStr.append("Abstract字体错、");
                    }
                } else if (themeFlag == 2) {
                    if (flag) {
                        flag = false;
                        abstractStr.append("-Abstract- 摘要内容字体错、");
                    } else {
                        abstractStr.append("摘要内容字体错、");
                    }
                } else if (themeFlag == 3) {
                    if (flag) {
                        flag = false;
                        abstractStr.append("-Abstract- Abstract和摘要内容字体全错、");
                    } else {
                        abstractStr.append("Abstract和摘要内容字体全错、");
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i)
                , engAbstractSize);
        if (!sizeFlag) {
            if (flag) {
                flag = false;
                abstractStr.append("-Abstract- 字体大小出错、");
            } else {
                abstractStr.append("字体大小出错、");
            }
        }
        if (flag == false) {
            abstractStr.deleteCharAt(abstractStr.length() - 1);
            abstractStr.append("!\n");
        }
        return abstractStr;
    }

    //中文关键词部分
    public StringBuilder chnKeywordsPart(int i) {
        StringBuilder abstractStr = new StringBuilder();
        boolean flag = true;
        if (!paragraphs.get(i - 1).getText().trim().equals("")) {
            flag = false;
            abstractStr.append("-关键词- 关键词前未空行、");
        }
        try {
            if (pf.getParaFirstLineChars(document, paragraphs.get(i)) != 0) {
                if (flag) {
                    flag = false;
                    abstractStr.append("-关键词- 关键词未顶格、");
                } else {
                    abstractStr.append("关键词未顶格、");
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        int themeFlag = this.isCorrectAbstractTheme1(document
                , paragraphs.get(i), chnKeywordsHeaderTheme
                , chnKeywordsTextTheme);
        if (themeFlag == 1) {
            if (flag) {
                flag = false;
                abstractStr.append("-关键词- 关键词字体错、");
            } else {
                abstractStr.append("关键词字体错、");
            }
        } else if (themeFlag == 2) {
            if (flag) {
                flag = false;
                abstractStr.append("-关键词- 关键词内容字体错、");
            } else {
                abstractStr.append("关键词内容字体错、");
            }
        } else if (themeFlag == 3) {
            if (flag) {
                flag = false;
                abstractStr.append("-关键词- 关键词和关键词内容字体全错、");
            } else {
                abstractStr.append("关键词和关键词内容字体全错、");
            }
        }
        boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i)
                , chnKeywordsSize);
        if (!sizeFlag) {
            if (flag) {
                flag = false;
                abstractStr.append("-关键词- 字体大小出错、");
            } else {
                abstractStr.append("字体大小出错、");
            }
        }
        if (flag == false) {
            abstractStr.deleteCharAt(abstractStr.length() - 1);
            abstractStr.append("!\n");
        }
        return abstractStr;
    }

    //英文关键词部分
    public StringBuilder engKeywordsPart(int i) {
        StringBuilder abstractStr = new StringBuilder();
        boolean flag = true;
        if (!paragraphs.get(i - 1).getText().trim().equals("")) {
            flag = false;
            abstractStr.append("-Keywords- Keywords前未空行、");
        }
        try {
            if (pf.getParaFirstLineChars(document, paragraphs.get(i)) != 0) {
                if (flag) {
                    flag = false;
                    abstractStr.append("-Keywords- Keywords未顶格、");
                } else {
                    abstractStr.append("Keywords未顶格、");
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        try {
            if (!textFormat.getFontBold(document, paragraphs.get(i), paragraphs.get(i).getRuns().get(0))) {
                if (flag) {
                    flag = false;
                    abstractStr.append("-Keywords- Keywords未加粗、");
                } else {
                    abstractStr.append("Keywords未加粗、");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        int themeFlag = this.isCorrectAbstractTheme1(document
                , paragraphs.get(i), engKeywordsHeaderTheme
                , engKeywordsTextTheme);
        if (themeFlag == 1) {
            if (flag) {
                flag = false;
                abstractStr.append("-Keywords- Keywords字体错、");
            } else {
                abstractStr.append("Keywords字体错、");
            }
        } else if (themeFlag == 2) {
            if (flag) {
                flag = false;
                abstractStr.append("-Keywords- 内容字体错、");
            } else {
                abstractStr.append("内容字体错、");
            }
        } else if (themeFlag == 3) {
            if (flag) {
                flag = false;
                abstractStr.append("-Keywords- Keywords和内容字体全错、");
            } else {
                abstractStr.append("Keywords和内容字体全错、");
            }
        }
        boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i)
                , engKeywordsSize);
        if (!sizeFlag) {
            if (flag) {
                flag = false;
                abstractStr.append("-Keywords- 字体大小出错、");
            } else {
                abstractStr.append("字体大小出错、");
            }
        }
        if (flag == false) {
            abstractStr.deleteCharAt(abstractStr.length() - 1);
            abstractStr.append("!\n");
        }
        return abstractStr;
    }

    //目录和致谢部分
    public StringBuilder cataPart(int i, int mar, String str) {
        StringBuilder textStr = new StringBuilder();
        boolean centerFlag = this.isCenter(document, paragraphs.get(i));
        boolean themeFlag = this.
                isCorrectTheme(document, paragraphs.get(i), cataTheme);
        //二号
        boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i)
                , cataSize);
        boolean marFlag = true;
        try {
//            System.out.println(pf.getParaSpacing(paragraphs.get(i), document) + str);
            marFlag = (pf.getParaSpacing(paragraphs.get(i), document) == mar) ? true : false;
//            System.out.println(marFlag);
        } catch (Exception e) {
            e.printStackTrace();
        }
        if (!themeFlag || !sizeFlag || !centerFlag || !marFlag) {
            textStr.append("-" + str + "- ");
            if (!themeFlag) {
                textStr.append("字体样式出错、");
            }
            if (!sizeFlag) {
                textStr.append("字体大小出错、");
            }
            if (!centerFlag) {
                textStr.append("未居中");
            }
            if (!marFlag) {
                textStr.append("行距错误");
            }
            if (textStr.toString().endsWith("、")) {
                textStr.deleteCharAt(textStr.length() - 1);
                textStr.append("!\n");
            } else {
                textStr.append("!\n");
            }
        }
        return textStr;
    }

    //目录
    public StringBuilder cataTextPart(int i) {
        StringBuilder textStr = new StringBuilder();
        boolean themeFlag = true;
        boolean sizeFlag = true;
        try {
            for (int j = 0; j < paragraphs.get(i).getCTP().getHyperlinkList().get(0).getRList().size(); j++) {
                if (paragraphs.get(i).getCTP().getHyperlinkList().get(0).getRList().get(j).getRPr()
                        .getRFonts() != null) {
                    if (paragraphs.get(i).getCTP().getHyperlinkList().get(0).getRList().get(j).getRPr()
                            .getRFonts().getAscii() != null && !paragraphs.get(i).getCTP().getHyperlinkList().get(0).getRList().get(j).getRPr()
                            .getRFonts().getAscii().equals(cataTextTheme)) themeFlag = false;
                    break;
                } else if (paragraphs.get(i).getStyleID() != null) {
                    if (document.getStyles().getStyle(paragraphs.get(i).getStyleID()) != null) {
                        if (document.getStyles().getStyle(paragraphs.get(i).getStyleID()).getCTStyle().getRPr() != null) {
                            if (document.getStyles().getStyle(paragraphs.get(i).getStyleID()).getCTStyle().getRPr().getRFonts() != null) {
                                TextType ty = new TextType();
                                CTFonts rFonts1 = document.getStyles().getStyle(paragraphs.get(i).getStyleID()).getCTStyle().getRPr().getRFonts();
                                if (rFonts1.getAscii() != null && rFonts1.getAscii().equals(cataTextTheme)) {
                                    themeFlag = true;
                                } else {
                                    themeFlag = false;
                                }
                            }
                        }
                    }
                } else {
                    themeFlag = false;
                    break;
                }
            }
            for (int j = 0; j < paragraphs.get(i).getCTP().getHyperlinkList().get(0).getRList().size(); j++) {
                if (paragraphs.get(i).getCTP().getHyperlinkList().get(0).getRList().get(j).getRPr() != null
                        && paragraphs.get(i).getCTP().getHyperlinkList().get(0).getRList().get(j).getRPr()
                        .getSz() != null) {
                    if (paragraphs.get(i).getCTP().getHyperlinkList().get(0).getRList().get(j).getRPr()
                            .getSz() != null && (paragraphs.get(i).getCTP().getHyperlinkList().get(0).getRList().get(j).getRPr()
                            .getSz().getVal().longValue()) / 2 != cataTextSize) sizeFlag = false;
                    break;
                } else if (paragraphs.get(i).getStyleID() != null) {
                    if (document.getStyles().getStyle(paragraphs.get(i).getStyleID()) != null) {
                        if (document.getStyles().getStyle(paragraphs.get(i).getStyleID()).getCTStyle().getRPr() != null) {
                            if (document.getStyles().getStyle(paragraphs.get(i).getStyleID()).getCTStyle().getRPr().getSz() != null) {
                                float rFonts1 = (float) document.getStyles().getStyle(paragraphs.get(i).getStyleID()).getCTStyle().getRPr().getSz().getVal().longValue() / 2;
                                if (rFonts1 == cataTextSize) {
                                    themeFlag = true;
                                } else {
                                    themeFlag = false;
                                }
                            }
                        }
                    }
                } else {
                    sizeFlag = false;
                    break;
                }
            }
        } catch (Exception e) {
            System.out.println(paragraphs.get(i).getText() + e);
            e.printStackTrace();
        }
        boolean marFlag = true;
        try {
            marFlag = pf.getParaSpacing(paragraphs.get(i), document) == 360;
        } catch (Exception e) {
            e.printStackTrace();
        }
        if (!themeFlag || !sizeFlag || !marFlag) {
            textStr.append("-目录内容- [" + (paragraphs.get(i).getText().length() > 6 ?
                    paragraphs.get(i).getText().substring(0, 6) : paragraphs.get(i).getText())
                    + "...] ");
            if (!themeFlag) {
                textStr.append("字体样式出错、");
            }
            if (!sizeFlag) {
                textStr.append("字体大小出错、");
            }
            if (!marFlag) {
                textStr.append("行距错误");
            }
            if (textStr.toString().endsWith("、")) {
                textStr.deleteCharAt(textStr.length() - 1);
                textStr.append("!\n");
            } else {
                textStr.append("!\n");
            }
        }
        return textStr;
    }

    //参考文献
    public StringBuilder litPart(int i) {
        StringBuilder textStr = new StringBuilder();
        boolean flag = true;
        try {
            if (!textFormat.getFontBold(document, paragraphs.get(i), paragraphs.get(i).getRuns().get(0))) {
                if (flag) {
                    flag = false;
                    textStr.append("-参考文献- 未加粗、");
                } else {
                    textStr.append("未加粗、");
                }
            }
            if (pf.getParaFirstLineChars(document, paragraphs.get(i)) != 0) {
                if (flag) {
                    flag = false;
                    textStr.append("-参考文献- 未顶格、");
                } else {
                    textStr.append("未顶格、");
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        boolean marFlag = true;
        try {
            marFlag = pf.getParaSpacing(paragraphs.get(i), document) == 360;
        } catch (Exception e) {
            e.printStackTrace();
        }
        boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i), litSize);
        boolean themeFlag = this.isCorrectTheme(document, paragraphs.get(i), litTheme);
        if (!themeFlag || !sizeFlag || !marFlag) {
            if (flag) {
                flag = false;
                textStr.append("-参考文献- ");
            }
            if (!themeFlag) {
                textStr.append("字体样式出错、");
            }
            if (!sizeFlag) {
                textStr.append("字体大小出错、");
            }
            if (!marFlag) {
                textStr.append("行距错误、");
            }
        }
        if (!flag && textStr.toString().endsWith("、")) {
            textStr.deleteCharAt(textStr.length() - 1);
            textStr.append("!\n");
        }
        return textStr;
    }

    //参考文献内容
    public StringBuilder litTextPart(int i) {
        StringBuilder textStr = new StringBuilder();
        boolean flag = true;
        try {
            if (pf.getParaFirstLineChars(document, paragraphs.get(i)) != 0) {
                if (flag) {
                    flag = false;
                    textStr.append("-参考文献内容- 未顶格、");
                } else {
                    textStr.append("未顶格、");
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        boolean marFlag = true;
        try {
            marFlag = pf.getParaSpacing(paragraphs.get(i), document) == 360;
        } catch (Exception e) {
            e.printStackTrace();
        }
        boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i), litTextSize);
        boolean themeFlag = true;
        if (!this.isCorrectTheme(document, paragraphs.get(i), chnlitTextTheme) &&
                !this.isCorrectTheme(document, paragraphs.get(i), englitTextTheme)) {
            themeFlag = false;
        }
        if (!themeFlag || !sizeFlag || !marFlag) {
            if (flag) {
                flag = false;
                textStr.append("-参考文献内容- [" + (paragraphs.get(i).getText().length() > 6 ?
                        paragraphs.get(i).getText().substring(0, 6) : paragraphs.get(i).getText())
                        + "...]");
            }
            if (!themeFlag) {
                textStr.append("字体样式出错、");
            }
            if (!sizeFlag) {
                textStr.append("字体大小出错、");
            }
            if (!marFlag) {
                textStr.append("行距错误、");
            }
        }
        if (!flag && textStr.toString().endsWith("、")) {
            textStr.deleteCharAt(textStr.length() - 1);
            textStr.append("!\n");
        }
        return textStr;
    }

    //正文标题部分1，2级
    public StringBuilder textHeaderPart1(int i, int val, String po, int mar) {
        StringBuilder headerStr = new StringBuilder();
        boolean centerFlag = true;
        if (po.equals("center")) {
            centerFlag = this.isCenter(document, paragraphs.get(i));
        } else {
            centerFlag = this.isLeft(document, paragraphs.get(i));
        }
        boolean marFlag = true;
        try {
            marFlag = pf.getParaSpacing(paragraphs.get(i), document) == mar;
        } catch (Exception e) {
            System.out.println(paragraphs.get(i).getText() + e);
            e.printStackTrace();
        }
        boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i), textHeader1Size);
        boolean themeFlag = this.isCorrectTheme(document, paragraphs.get(i), textHeader1Theme);
        if (!themeFlag || !sizeFlag || !centerFlag || !marFlag) {
            headerStr.append("-" + val + "级标题- [" + (paragraphs.get(i).getText().length() > 6 ?
                    paragraphs.get(i).getText().substring(0, 6) : paragraphs.get(i).getText())
                    + "...] ");
            if (!themeFlag) {
                headerStr.append("字体样式出错、");
            }
            if (!sizeFlag) {
                headerStr.append("字体大小出错、");
            }
            if (!centerFlag) {
                if (po.equals("center")) {
                    headerStr.append("未居中、");
                } else {
                    headerStr.append("未居左、");
                }
            }
            if (!marFlag) {
                headerStr.append("行距错误、");
            }
            if (headerStr.toString().endsWith("、")) {
                headerStr.deleteCharAt(headerStr.length() - 1);
                headerStr.append("!\n");
            } else {
                headerStr.append("!\n");
            }
        }
        return headerStr;
    }

    //正文标题部分3级
    public StringBuilder textHeaderPart2(int i, int val) {
        StringBuilder headerStr = new StringBuilder();
        boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i), textHeader2Size);
        boolean themeFlag = this.isCorrectTheme(document, paragraphs.get(i), textHeader2Theme);
        if (!themeFlag || !sizeFlag) {
            if (!sizeFlag && !themeFlag) {
                headerStr.append("-" + val + "级标题- [" + (paragraphs.get(i).getText().length() > 6 ?
                        paragraphs.get(i).getText().substring(0, 6) : paragraphs.get(i).getText())
                        + "...] 段落字体大小及样式出错!\n");
            } else if (!themeFlag) {
                headerStr.append("-" + val + "级标题- [" + (paragraphs.get(i).getText().length() > 6 ?
                        paragraphs.get(i).getText().substring(0, 6) : paragraphs.get(i).getText())
                        + "...] 段落字体样式出错!\n");
            } else if (!sizeFlag) {
                headerStr.append("-" + val + "级标题- [" + (paragraphs.get(i).getText().length() > 6 ?
                        paragraphs.get(i).getText().substring(0, 6) : paragraphs.get(i).getText())
                        + "...] 段落字体大小出错!\n");
            }
        }
        return headerStr;
    }

    //正文
    public StringBuilder textPart(int i) {
        StringBuilder textStr = new StringBuilder();
        boolean flag = true;
        try {
            if (pf.getParaFirstLineChars(document, paragraphs.get(i)) != 200) {
                if (flag) {
                    flag = false;
                    textStr.append("-正文内容- [" + (paragraphs.get(i).getText().length() > 6 ?
                            paragraphs.get(i).getText().substring(0, 6) : paragraphs.get(i).getText())
                            + "...] 首行未缩进、");
                } else {
                    textStr.append("首行未缩进、");
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        boolean marFlag = true;
        try {
            marFlag = pf.getParaSpacing(paragraphs.get(i), document) == 360;
        } catch (Exception e) {
            e.printStackTrace();
        }
        boolean sizeFlag = this.isCorrectSize(document, paragraphs.get(i), textSize);
        boolean themeFlag = true;
        if (!this.isCorrectTheme(document, paragraphs.get(i), textTheme)) {
            themeFlag = false;
        }
        if (!themeFlag || !sizeFlag || !marFlag) {
            if (flag) {
                flag = false;
                textStr.append("-正文内容- [" + (paragraphs.get(i).getText().length() > 6 ?
                        paragraphs.get(i).getText().substring(0, 6) : paragraphs.get(i).getText())
                        + "...] ");
            }
            if (!themeFlag) {
                textStr.append("字体样式出错、");
            }
            if (!sizeFlag) {
                textStr.append("字体大小出错、");
            }
            if (!marFlag) {
                textStr.append("行距错误、");
            }
        }
        if (!flag && textStr.toString().endsWith("、")) {
            textStr.deleteCharAt(textStr.length() - 1);
            textStr.append("!\n");
        }
        return textStr;
    }

}
package com.mng.domain;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;

public class TextFormat {
    /**
     * 获取文档默认字体大小
     *
     * @param docx
     * @return
     * @throws Exception
     */
    public float getDocxDefaultFontSize(XWPFDocument docx) throws Exception {
        if (docx.getStyle().getDocDefaults().getRPrDefault() != null) {
            if (docx.getStyle().getDocDefaults().getRPrDefault().getRPr() != null) {
                CTRPr rPr = docx.getStyle().getDocDefaults().getRPrDefault().getRPr();
                if (rPr.getSz() != null) {
                    return (float) rPr.getSz().getVal().longValue() / 2;
                } else if (rPr.getSzCs() != null) {
                    return (float) rPr.getSzCs().getVal().longValue() / 2;
                }
            }
        }
        return 10;
    }

    /**
     * 获取字体大小
     *
     * @param docx
     * @param paragraph
     * @param x
     * @return
     * @throws Exception
     */
    public float getFontSize(XWPFDocument docx, XWPFParagraph paragraph, XWPFRun x) throws Exception {

        if (x.getCTR().getRPr() != null) {
            if (x.getCTR().getRPr().getSz() != null) {
                return (float) x.getCTR().getRPr().getSz().getVal().longValue() / 2;
            } else if (x.getCTR().getRPr().getSzCs() != null) {
                return (float) x.getCTR().getRPr().getSzCs().getVal().longValue() / 2;
            }
        }

        if (docx.getStyles().getStyle(paragraph.getStyleID()) != null) {
            if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getRPr() != null) {
                CTRPr rPr = docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getRPr();
                if (rPr.getSz() != null) return (float) rPr.getSz().getVal().longValue() / 2;
                else if (rPr.getSzCs() != null) return (float) rPr.getSzCs().getVal().longValue() / 2;
            }
            if (docx.getStyles().getStyle(paragraph.getStyleID()).getLinkStyleID() != null) {
                if (docx.getStyles().getStyle(docx.getStyles().getStyle(paragraph.getStyleID()).getLinkStyleID())
                        .getCTStyle().getRPr() != null) {
                    CTRPr rPr = docx.getStyles().getStyle(docx.getStyles().getStyle(paragraph.getStyleID())
                            .getLinkStyleID()).getCTStyle().getRPr();
                    if (rPr.getSz() != null) return (float) rPr.getSz().getVal().longValue() / 2;
                    else if (rPr.getSzCs() != null) return (float) rPr.getSzCs().getVal().longValue() / 2;
                }
            }
        }

        //文档默认字体大小
        float fontSize = this.getDocxDefaultFontSize(docx);
        return fontSize;
    }

    /**
     * 获取字体的字体主题
     *
     * @param run
     * @param docx
     * @param paragraph
     * @return
     * @throws Exception
     */
    public String getFontTheme(XWPFRun run, XWPFDocument docx, XWPFParagraph paragraph) throws Exception {
        if (run.getCTR().getRPr() != null && run.getCTR().getRPr().getRFonts() != null) {
            CTFonts rFonts = run.getCTR().getRPr().getRFonts();
            TextType ty = new TextType();
            if (ty.isEnglishFont(run) && rFonts.getAscii() != null) {
                return rFonts.getAscii();
            } else if (ty.isChinessFont(run) && rFonts.getEastAsia() != null) {
                return rFonts.getEastAsia();
            } else if (rFonts.getAscii() != null) {
                return rFonts.getAscii();
            }
        }

        if (docx.getStyles().getStyle(paragraph.getStyleID()) != null) {
            if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getRPr() != null) {
                if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getRPr().getRFonts() != null) {
                    TextType ty = new TextType();
                    CTFonts rFonts1 = docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getRPr().getRFonts();
                    if ( rFonts1.getAscii() != null) return rFonts1.getAscii();
                    else if ( rFonts1.getEastAsia() != null) return rFonts1.getEastAsia();
                }
            }
            if (docx.getStyles().getStyle(paragraph.getStyleID()).getLinkStyleID() != null) {
                if (docx.getStyles().getStyle(docx.getStyles().getStyle(paragraph.getStyleID())
                        .getLinkStyleID()).getCTStyle().getRPr() != null) {
                    if (docx.getStyles().getStyle(docx.getStyles().getStyle(paragraph.getStyleID())
                            .getLinkStyleID()).getCTStyle().getRPr().getRFonts() != null) {
                        TextType ty = new TextType();
                        CTFonts rFonts1 = docx.getStyles().getStyle(docx.getStyles().getStyle(paragraph.getStyleID())
                                .getLinkStyleID()).getCTStyle().getRPr().getRFonts();
                        if ( rFonts1.getAscii() != null) return rFonts1.getAscii();
                        else if ( rFonts1.getEastAsia() != null) return rFonts1.getEastAsia();
                    }
                }
            }
        }

        //默认字体主题
        String fontTheme = "";
        //如果为英文字体
        TextType ty = new TextType();
        if (ty.isEnglishFont(run)) {
            fontTheme = docx.getStyle().getDocDefaults().getRPrDefault().getRPr().getRFonts().getAscii();
        } else {
            fontTheme = docx.getStyle().getDocDefaults().getRPrDefault().getRPr().getRFonts().getEastAsia();
        }
        return fontTheme;
    }

    /**
     * 获取默认字体加粗
     *
     * @param docx
     * @return
     * @throws Exception
     */
    public Boolean getDocxDefaultFontBold(XWPFDocument docx) throws Exception {
        if (docx.getStyle().getDocDefaults().getRPrDefault() != null) {
            if (docx.getStyle().getDocDefaults().getRPrDefault().getRPr() != null) {
                CTRPr rPr = docx.getStyle().getDocDefaults().getRPrDefault().getRPr();
                if (rPr.getB() != null) {
                    if (rPr.getB().isSetVal()) return false;
                    else return true;
                } else if (rPr.getBCs() != null) {
                    if (rPr.getBCs().isSetVal()) return false;
                    else return true;
                }
            }
        }
        return false;
    }

    /**
     * 获取字体是否加粗
     *
     * @param docx
     * @param paragraph
     * @param run
     * @return
     * @throws Exception
     */
    public Boolean getFontBold(XWPFDocument docx, XWPFParagraph paragraph, XWPFRun run) throws Exception {
        if (run.getCTR().getRPr() != null) {
            if (run.getCTR().getRPr().getB() != null) {
                if (run.getCTR().getRPr().getB().isSetVal()) return false;
                else return true;
            }
        }

        if (docx.getStyles().getStyle(paragraph.getStyleID()) != null) {
            if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getRPr() != null) {
                CTRPr rPr = docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getRPr();
                if (rPr.getB() != null) {
                    if (rPr.getB().isSetVal()) return false;
                    else return true;
                } else if (rPr.getBCs() != null) {
                    if (rPr.getBCs().isSetVal()) return false;
                    else return true;
                }
            }
            if (docx.getStyles().getStyle(paragraph.getStyleID()).getLinkStyleID() != null) {
                if (docx.getStyles().getStyle(docx.getStyles().getStyle(paragraph.getStyleID()).getLinkStyleID())
                        .getCTStyle().getRPr() != null) {
                    CTRPr rPr = docx.getStyles().getStyle(docx.getStyles().getStyle(paragraph.getStyleID())
                            .getLinkStyleID()).getCTStyle().getRPr();
                    if (rPr.getB() != null) {
                        if (rPr.getB().isSetVal()) return false;
                        else return true;
                    } else if (rPr.getBCs() != null) {
                        if (rPr.getBCs().isSetVal()) return false;
                        else return true;
                    }
                }
            }
        }

        //默认字体加粗状态
        return this.getDocxDefaultFontBold(docx);
    }

}

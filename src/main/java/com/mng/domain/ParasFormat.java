package com.mng.domain;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;

public class ParasFormat {

    /**
     * 获取文档默认首行缩进
     *
     * @param docx
     * @return
     * @throws Exception
     */
    public Integer getDocxFirstLineChars(XWPFDocument docx) throws Exception {
        if (docx.getStyle().getDocDefaults().getPPrDefault() != null && docx.getStyle().getDocDefaults().getPPrDefault().getPPr() != null) {
            CTPPr pPr = docx.getStyle().getDocDefaults().getPPrDefault().getPPr();
            if (pPr.getInd() != null && pPr.getInd().isSetFirstLineChars()) {
                return pPr.getInd().getFirstLineChars().intValue();
            }
        }
        return 0;
    }

    /**
     * 获取段落首行缩进
     *
     * @param docx
     * @param paragraph
     * @return
     * @throws Exception
     */
    public Integer getParaFirstLineChars(XWPFDocument docx, XWPFParagraph paragraph) throws Exception {
        if (paragraph.getCTP().getPPr() != null) {
            if (paragraph.getCTP().getPPr().getInd() != null && paragraph.getCTP().getPPr().getInd().isSetFirstLineChars()) {
                return paragraph.getCTP().getPPr().getInd().getFirstLineChars().intValue();
            }
        }

        if (docx.getStyles().getStyle(paragraph.getStyleID()) != null) {
            if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr() != null) {
                CTPPr pPr = docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr();
                if (pPr.getInd() != null && pPr.getInd().isSetFirstLineChars()) {
                    return pPr.getInd().getFirstLineChars().intValue();
                }
            }
        }

        return this.getDocxFirstLineChars(docx);
    }

    /**
     * 获得段前距离
     *
     * @param paragraph
     * @param docx
     * @return
     * @throws Exception
     */
    public Integer getParaSpacingBeforeLines(XWPFParagraph paragraph, XWPFDocument docx) throws Exception {
        if (paragraph.getSpacingBeforeLines() != -1) {
            return paragraph.getSpacingBeforeLines();
        }

        if (docx.getStyles().getStyle(paragraph.getStyleID()) != null) {
            if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr() != null) {
                if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getSpacing() != null) {
                    if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getSpacing().getBeforeLines() != null) {
                        return docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getSpacing().getBeforeLines().intValue();
                    }
                }
            }
        }
        // 文档默认段前行距
        Integer beforeLines = this.getDocxDefaultSpacing(docx, "SPACING_BEFORE_Line");
        return beforeLines;
    }

    /**
     * 获得段前磅数
     *
     * @param paragraph
     * @param docx
     * @return
     * @throws Exception
     */
    public Integer getParaSpacingBefore(XWPFParagraph paragraph, XWPFDocument docx) throws Exception {
        if (paragraph.getSpacingBefore() != -1) {
            return paragraph.getSpacingBefore();
        }

        if (docx.getStyles().getStyle(paragraph.getStyleID()) != null) {
            if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr() != null) {
                if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getSpacing() != null) {
                    if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getSpacing().getBefore() != null) {
                        return docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getSpacing().getBefore().intValue();
                    }
                }
            }
        }

        // 文档默认段前行距
        Integer spacingBefore = this.getDocxDefaultSpacing(docx, "SPACING_BEFORE");
        return spacingBefore;
    }

    /**
     * 获得段后距离
     *
     * @param paragraph
     * @param docx
     * @return
     * @throws Exception
     */
    public Integer getParaSpacingAfterLines(XWPFParagraph paragraph, XWPFDocument docx) throws Exception {
        if (paragraph.getSpacingAfterLines() != -1) {
            return paragraph.getSpacingAfterLines();
        }

        if (docx.getStyles().getStyle(paragraph.getStyleID()) != null) {
            if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr() != null) {
                if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getSpacing() != null) {
                    if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getSpacing().getAfterLines() != null) {
                        return docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getSpacing().getAfterLines().intValue();
                    }
                }
            }
        }
        // 文档默认段后行距
        Integer afterLines = this.getDocxDefaultSpacing(docx, "SPACING_AFTER_LINE");
        return afterLines;
    }

    /**
     * 获得段后磅数
     *
     * @param paragraph
     * @param docx
     * @return
     * @throws Exception
     */
    public Integer getParaSpacingAfter(XWPFParagraph paragraph, XWPFDocument docx) throws Exception {
        if (paragraph.getSpacingAfter() != -1) {
            return paragraph.getSpacingAfter();
        }

        if (docx.getStyles().getStyle(paragraph.getStyleID()) != null) {
            if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr() != null) {
                if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getSpacing() != null) {
                    if (docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getSpacing().getAfter() != null) {
                        return docx.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr().getSpacing().getAfter().intValue();
                    }
                }
            }
        }

        // 文档默认段后行距
        Integer spacingAfter = this.getDocxDefaultSpacing(docx, "SPACING_AFTER");
        return spacingAfter;
    }

    public Integer getParaSpacing(XWPFParagraph paragraph, XWPFDocument document) throws Exception {
        if (paragraph.getCTP().getPPr() != null) {
            if (paragraph.getCTP().getPPr().getSpacing() != null) {
                return Integer.parseInt(paragraph.getCTP().getPPr().getSpacing().getLine().toString());
            }
        }

        if (document.getStyles().getStyle(paragraph.getStyleID()) != null) {
            if (document.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr() != null) {
                CTPPr pPr = document.getStyles().getStyle(paragraph.getStyleID()).getCTStyle().getPPr();
                if (pPr.getSpacing() != null && pPr.getSpacing().getLine() != null)
                    return Integer.parseInt(pPr.getSpacing().getLine().toString());
            }
            if (document.getStyles().getStyle(paragraph.getStyleID()).getLinkStyleID() != null) {
                if (document.getStyles().getStyle(document.getStyles().getStyle(paragraph.getStyleID()).getLinkStyleID())
                        .getCTStyle().getPPr() != null) {
                    CTPPr pPr = document.getStyles().getStyle(document.getStyles().getStyle(paragraph.getStyleID())
                            .getLinkStyleID()).getCTStyle().getPPr();
                    if (pPr.getSpacing() != null && pPr.getSpacing().getLine() != null)
                        return Integer.parseInt(pPr.getSpacing().getLine().toString());
                }
            }
        }

        return -1;
    }

    /**
     * 获取文档默认段前、后行距
     *
     * @param docx
     * @param category
     * @return
     * @throws Exception
     */
    public Integer getDocxDefaultSpacing(XWPFDocument docx, String category) throws Exception {
        if (docx.getStyle().getDocDefaults().getPPrDefault().getPPr() != null) {
            if (docx.getStyle().getDocDefaults().getPPrDefault().getPPr().getSpacing() != null) {
                CTSpacing spacing = docx.getStyle().getDocDefaults().getPPrDefault().getPPr().getSpacing();
                if (category.equals("SPACING_BEFORE")) {
                    if (spacing.getBefore() != null) {
                        return spacing.getBefore().intValue();
                    }
                } else if (category.equals("SPACING_BEFORE_Line")) {
                    if (spacing.getBeforeLines() != null) {
                        return spacing.getBeforeLines().intValue();
                    }
                } else if (category.equals("SPACING_AFTER")) {
                    if (spacing.getAfter() != null) {
                        return spacing.getAfter().intValue();
                    }
                } else if (category.equals("SPACING_AFTER_LINE")) {
                    if (spacing.getAfterLines() != null) {
                        return spacing.getAfterLines().intValue();
                    }
                }
            }
        }
        return 0;
    }
}

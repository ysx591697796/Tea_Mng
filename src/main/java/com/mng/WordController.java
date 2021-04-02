package com.mng;

import com.mng.application.TextFormatApplication;
import com.mng.domain.TextFormat;
import com.mng.domain.WordToPoi;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

@RestController
@RequestMapping("api")
public class WordController {

    @RequestMapping("/fileupload")
    public String uploadfile(@RequestParam("file") MultipartFile request) throws IOException {
        if (request == null) {
            return "fail";
        }

        //将word转换为XWPFDocument
        XWPFDocument document = new XWPFDocument(request.getInputStream());
        //XWPFParagraph
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        //WordToPoi
        WordToPoi wp = new WordToPoi();
        //TextFormat
        TextFormat tf = new TextFormat();
        //TextFormatApplication
        TextFormatApplication textFormatApplication = new TextFormatApplication();

        //Page错误
        StringBuilder pageStr = new StringBuilder();
        //Mar错误
        StringBuilder marStr = new StringBuilder();
        //Header错误
        StringBuilder headerStr = new StringBuilder();

        //页大小 A4:11906W,16838H
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        CTPageSz pageSz = sectPr.getPgSz();
        if (!pageSz.getW().toString().equals("11906") || !pageSz.getH().toString().equals("16838")) {
            headerStr.append("-页面大小- 页面大小出错，需设置A4!\n");
        } else {

        }
//        System.out.println(pageSz.getW() + " " + pageSz.getH() + " " + pageSz.getOrient());

        //页边距 top:1701(30mm) bottom:1418(25mm) left:1701(30mm) right:1134(20mm)
        CTPageMar ctPageMar = sectPr.getPgMar();
        if (!ctPageMar.getTop().toString().equals("1701") || !ctPageMar.getBottom().toString().equals("1418") ||
                !ctPageMar.getLeft().toString().equals("1701") || !ctPageMar.getRight().toString().equals("1134")) {
            marStr.append("-页面边距- 页面边距出错，");
            if (!ctPageMar.getTop().toString().equals("1701")) {
                marStr.append("上(需设置30mm)");
            } else if (!ctPageMar.getBottom().toString().equals("1418")) {
                marStr.append("下(需设置25mm)");
            } else if (!ctPageMar.getLeft().toString().equals("1701")) {
                marStr.append("左(需设置30mm)");
            } else if (!ctPageMar.getRight().toString().equals("1134")) {
                marStr.append("有(需设置20mm)");
            } else {

            }
            marStr.append("!\n");
        }
//        System.out.println("top:" + ctPageMar.getTop() + " bottom:" + ctPageMar.getBottom()
//                + " left:" + ctPageMar.getLeft() + " right:" + ctPageMar.getRight());

        //正文
        if (paragraphs.size() != 0) {
            for (int i = 0; i < paragraphs.size(); i++) {
                if (wp.getTitleLvl(document, paragraphs.get(i)).equals("0")) {
//                    boolean flag = true;
                    boolean sizeFlag = textFormatApplication.isCorrectSize(document, paragraphs.get(i), (float) 12.0);
                    boolean themeFlag = textFormatApplication.isCorrectTheme(document, paragraphs.get(i), "黑体");
//                    List<XWPFRun> runs = paragraphs.get(i).getRuns();
//                    for (XWPFRun r : runs) {
//                        if (!r.getText(0).equals(" ")) {
//                            try {
//                                if (!tf.getFontTheme(r, document, paragraphs.get(i)).equals("黑体")) {
//                                    flag = false;
//                                    themeFlag = false;
//                                }
//                                if (tf.getFontSize(document, paragraphs.get(i), r) != 10.5) {
//                                    flag = false;
//                                    sizeFlag = false;
//                                }
//                            } catch (Exception e) {
//                                e.printStackTrace();
//                            }
//                        }
//                    }
                    if (!themeFlag || !sizeFlag) {
                        if (!sizeFlag && !themeFlag) {
                            headerStr.append("-1级标题- [" + paragraphs.get(i).getText().substring(0, 6)
                                    + "...] 段落字体大小及样式出错!\n");
                        } else if (!themeFlag) {
                            headerStr.append("-1级标题- [" + paragraphs.get(i).getText().substring(0, 6)
                                    + "...] 段落字体样式出错!\n");
                        } else if (!sizeFlag) {
                            headerStr.append("-1级标题- [" + paragraphs.get(i).getText().substring(0, 6)
                                    + "...] 段落字体大小出错!\n");
                        }
                    } else {
//                        System.out.println("-1级标题- [" + paragraphs.get(i).getText().substring(0, 6)
//                                + "...] 段落无误!");
                    }
                } else if (wp.getTitleLvl(document, paragraphs.get(i)).equals("1")) {
                    boolean sizeFlag = textFormatApplication.isCorrectSize(document, paragraphs.get(i), (float) 12.0);
                    boolean themeFlag = textFormatApplication.isCorrectTheme(document, paragraphs.get(i), "黑体");
//                    List<XWPFRun> runs = paragraphs.get(i).getRuns();
//                    for (XWPFRun r : runs) {
//                        if (!r.getText(0).equals(" ")) {
//                            try {
//                                if (!tf.getFontTheme(r, document, paragraphs.get(i)).equals("黑体") ||
//                                        tf.getFontTheme(r, document, paragraphs.get(i)) == null) {
//                                    flag = false;
//                                    themeFlag = false;
//                                }
//                                if (tf.getFontSize(document, paragraphs.get(i), r) != 10.5) {
//                                    flag = false;
//                                    sizeFlag = false;
//                                }
//                            } catch (Exception e) {
//                                e.printStackTrace();
//                            }
//                        }
//                    }
                    if (!themeFlag || !sizeFlag) {
                        if (!sizeFlag && !themeFlag) {
                            headerStr.append("-2级标题- [" + paragraphs.get(i).getText().substring(0, 6)
                                    + "...] 段落字体大小及样式出错!\n");
                        } else if (!themeFlag) {
                            headerStr.append("-2级标题- [" + paragraphs.get(i).getText().substring(0, 6)
                                    + "...] 段落字体样式出错!\n");
                        } else if (!sizeFlag) {
                            headerStr.append("-2级标题- [" + paragraphs.get(i).getText().substring(0, 6)
                                    + "...] 段落字体大小出错!\n");
                        }
                    } else {
//                        System.out.println("-2级标题- [" + paragraphs.get(i).getText().substring(0, 6)
//                                + "...] 段落无误!");
                    }
                } else {
//                    List<XWPFRun> runs = paragraphs.get(i).getRuns();
//                    for (XWPFRun r : runs) {
//                        try {
//                            System.out.println(tf.getFontBold(document,paragraphs.get(i),r));
//                        } catch (Exception e) {
//                            e.printStackTrace();
//                        }
//                    }
                }
            }
        }

        if (pageStr != null && pageStr.length() > 0) {
            System.out.println("---页面部分---:\n" + pageStr);
        } else {
            System.out.println("---页面部分---:（无错）");
        }
        if (marStr != null && marStr.length() > 0) {
            System.out.println("---页边距部分---:\n" + pageStr);
        } else {
            System.out.println("---页边距部分---:（无错）");
        }
        if (headerStr != null && headerStr.length() > 0) {
            System.out.println("---标题部分---:\n" + headerStr);
        } else {
            System.out.println("---标题部分---:（无错）");
        }

        return "success";
    }
}

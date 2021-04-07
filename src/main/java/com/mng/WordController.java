package com.mng;

import com.mng.application.TextFormatApplication;
import com.mng.domain.WordToPoi;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("api")
public class WordController {

    @RequestMapping("/fileupload")
    public Map<String, String> uploadfile(@RequestParam("file") MultipartFile request) throws IOException {
        Map<String, String> tag = new HashMap<>();
        if (request == null) {
            tag.put("status", "fail");
            return tag;
        }

        //将word转换为XWPFDocument
        XWPFDocument document = new XWPFDocument(request.getInputStream());
        //XWPFParagraph
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        //WordToPoi
        WordToPoi wp = new WordToPoi();
        //TextFormatApplication
        TextFormatApplication textFormatApplication = new TextFormatApplication(document);

        //Page错误
        String pageStr = textFormatApplication.pageSizePart();
        //Mar错误
        String marStr = textFormatApplication.pageMarPart();
        //Footer错误
        String footerStr = textFormatApplication.footerPart();
        //PaperHeader错误
        StringBuilder paperHStr = new StringBuilder();
        //Abstract错误
        StringBuilder abstractStr = new StringBuilder();
        //Header错误
        StringBuilder headerStr = new StringBuilder();
        //Text错误
        StringBuilder textStr = new StringBuilder();
        //Text错误
        StringBuilder errStr = new StringBuilder();

        //Begin
        int begin = 0;
        //End
        int end = 0;

        if (paragraphs.size() != 0) {
            for (int i = 0; i < paragraphs.size(); i++) {
                if (wp.getTitleLvl(document, paragraphs.get(i)).equals("0") &&
                        paragraphs.get(i).getText().startsWith("第")) {
                    begin = i;
                    continue;
                }
                if (paragraphs.get(i).getText().trim().startsWith("致  谢") ||
                        paragraphs.get(i).getText().trim().startsWith("致 谢") ||
                        paragraphs.get(i).getText().trim().startsWith("致谢")) {
                    end = i;
                    continue;
                }
            }
        }

        //文章
        if (paragraphs.size() != 0) {
            for (int i = 0; i < paragraphs.size(); i++) {
                //专业段(中文)
                if (paragraphs.get(i).getText().startsWith("专业")) {
                    paperHStr.append(textFormatApplication.chnHeaderPart(i));
                    continue;
                }

                //专业段(英文)
                if (paragraphs.get(i).getText().startsWith("Major")) {
                    paperHStr.append(textFormatApplication.engHeaderPart(i));
                    continue;
                }

                //摘要段(中文)
                if (paragraphs.get(i).getText().trim().startsWith("摘 要") || paragraphs.get(i)
                        .getText().trim().startsWith("摘要")) {
                    abstractStr.append(textFormatApplication.chnAbstractpart(i));
                    continue;
                }

                //摘要段(英文)
                if (paragraphs.get(i).getText().trim().startsWith("Abstract")) {
                    abstractStr.append(textFormatApplication.engAbstractpart(i));
                    continue;
                }

                //关键词
                if (paragraphs.get(i).getText().trim().startsWith("关键词")) {
                    abstractStr.append(textFormatApplication.chnKeywordsPart(i));
                    continue;
                }

                //关键词(英文)
                if (paragraphs.get(i).getText().trim().startsWith("Key words")) {
                    abstractStr.append(textFormatApplication.engKeywordsPart(i));
                    continue;
                }

                //目录
                if (paragraphs.get(i).getText().trim().startsWith("目  录") ||
                        paragraphs.get(i).getText().trim().startsWith("目 录") ||
                        paragraphs.get(i).getText().trim().startsWith("目录")) {
                    textStr.append(textFormatApplication.cataPart(i, 720, "目录"));
                    continue;
                }

                //正文
                if (wp.getTitleLvl(document, paragraphs.get(i)).equals("0")) {
                    headerStr.append(textFormatApplication.textHeaderPart1(i, 1, "center", 720));
                    continue;
                } else if (wp.getTitleLvl(document, paragraphs.get(i)).equals("1")) {
                    headerStr.append(textFormatApplication.textHeaderPart1(i, 2, "left", 360));
                    continue;
                } else if (wp.getTitleLvl(document, paragraphs.get(i)).equals("2")) {
                    headerStr.append(textFormatApplication.textHeaderPart2(i, 3));
                    continue;
                } else if (wp.getTitleLvl(document, paragraphs.get(i)).startsWith("TOC")) {
                    textStr.append(textFormatApplication.cataTextPart(i));
                    continue;
                }

                //参考文献
                if (paragraphs.get(i).getText().trim().startsWith("参考文献")) {
                    textStr.append(textFormatApplication.litPart(i));
                    continue;
                }

                if (paragraphs.get(i).getText().trim().startsWith("[")) {
                    textStr.append(textFormatApplication.litTextPart(i));
                    continue;
                }

                //致谢
                if (paragraphs.get(i).getText().trim().startsWith("致  谢") ||
                        paragraphs.get(i).getText().trim().startsWith("致 谢") ||
                        paragraphs.get(i).getText().trim().startsWith("致谢")) {
                    textStr.append(textFormatApplication.cataPart(i, -1, "致谢"));
                    continue;
                }

            }
        }

        if (paragraphs.size() != 0) {
            for (int i = begin - 1; i < end; i++) {
                if (wp.getTitleLvl(document, paragraphs.get(i)).equals("text")
                        && paragraphs.get(i).getText() != null && !paragraphs.get(i).getText().trim().equals("")) {
                    textStr.append(textFormatApplication.textPart(i));
                }
            }
        }

        if (pageStr != null && pageStr.length() > 0) {
            errStr.append("---页面部分---:\n" + pageStr + "\n");
        } else {
            errStr.append("---页面部分---:（无错）\n\n");
        }
        tag.put("pageStr",pageStr.toString());
        if (marStr != null && marStr.length() > 0) {
            errStr.append("---页边距部分---:\n" + marStr + "\n");
        } else {
            errStr.append("---页边距部分---:（无错）\n\n");
        }
        tag.put("marStr",marStr.toString());
        if (footerStr != null && footerStr.length() > 0) {
            errStr.append("---页码部分---:\n" + footerStr + "\n");
        } else {
            errStr.append("---页码部分---:（无错）\n\n");
        }
        tag.put("footerStr",footerStr.toString());
        if (paperHStr != null && paperHStr.length() > 0) {
            errStr.append("---论文标题部分---:\n" + paperHStr + "\n");
        } else {
            errStr.append("---论文标题部分---:（无错）\n\n");
        }
        tag.put("paperHStr",paperHStr.toString());
        if (abstractStr != null && abstractStr.length() > 0) {
            errStr.append("---摘要部分---:\n" + abstractStr + "\n");
        } else {
            errStr.append("---摘要部分---:（无错）\n\n");
        }
        tag.put("abstractStr",abstractStr.toString());
        if (headerStr != null && headerStr.length() > 0) {
            errStr.append("---大纲标题部分---:\n" + headerStr + "\n");
        } else {
            errStr.append("---大纲标题部分---:（无错）\n\n");
        }
        tag.put("headerStr",headerStr.toString());
        if (textStr != null && textStr.length() > 0) {
            errStr.append("---正文部分---:\n" + textStr + "\n");
        } else {
            errStr.append("---正文部分---:（无错）\n\n");
        }
        tag.put("textStr",textStr.toString());

        SimpleDateFormat df = new SimpleDateFormat("HH-mm-ss");
        String path = "./" + df.format(new Date()) + ".doc";
        String name = df.format(new Date()) + ".doc";
        try {
            InputStream is = new ByteArrayInputStream(errStr.toString().getBytes("utf-8"));
            OutputStream os = new FileOutputStream(path);
            POIFSFileSystem fs = new POIFSFileSystem();
            fs.createDocument(is, "WordDocument");
            fs.writeFilesystem(os);
            fs.close();
            is.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        tag.put("status", "ok");
        tag.put("errFile", name);
        tag.put("errStr", errStr.toString());

        return tag;
    }

    @RequestMapping("/fileupload1")
    public void uploadfile1(@RequestParam("name") String name, HttpServletResponse response, HttpServletRequest request1) throws IOException {
        response.setCharacterEncoding("UTF-8"); //字符编码
        response.setContentType("multipart/form-data"); //二进制传输数据
        //设置响应头
        response.setHeader("Content-Disposition",
                "attachment;fileName=" + URLEncoder.encode(name, "UTF-8"));

        File file = new File("./" + name);
        //2、 读取文件--输入流
        InputStream input = new FileInputStream(file);
        //3、 写出文件--输出流
        OutputStream out = response.getOutputStream();

        byte[] buff = new byte[1024];
        int index = 0;
        //4、执行 写出操作
        while ((index = input.read(buff)) != -1) {
            out.write(buff, 0, index);
            out.flush();
        }
        out.close();
        input.close();
    }
}

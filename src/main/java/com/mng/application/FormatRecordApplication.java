package com.mng.application;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

public class FormatRecordApplication {
    MultipartFile file = null;

    public FormatRecordApplication(MultipartFile file) {
        this.file = file;
    }

    public void record() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        XSSFSheet sheet = workbook.getSheetAt(0);

        if (sheet.getRow(0).getCell(1).toString().trim() != null) {
            TextFormatApplication.chnHeaderTheme = sheet.getRow(0).getCell(1).toString();
        }
        if (sheet.getRow(1).getCell(1).toString().trim() != null) {
            TextFormatApplication.chnHeaderSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(1).getCell(1).toString());
        }
        if (sheet.getRow(2).getCell(1).toString().trim() != null) {
            TextFormatApplication.chnNameTheme = sheet.getRow(2).getCell(1).toString();
        }
        if (sheet.getRow(3).getCell(1).toString().trim() != null) {
            TextFormatApplication.chnNameSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(3).getCell(1).toString());
        }
        if (sheet.getRow(4).getCell(1).toString().trim() != null) {
            TextFormatApplication.chnMajorTheme = sheet.getRow(4).getCell(1).toString();
        }
        if (sheet.getRow(5).getCell(1).toString().trim() != null) {
            TextFormatApplication.chnMajorSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(5).getCell(1).toString());
        }
        if (sheet.getRow(6).getCell(1).toString().trim() != null) {
            TextFormatApplication.engHeaderTheme = sheet.getRow(6).getCell(1).toString();
        }
        if (sheet.getRow(7).getCell(1).toString().trim() != null) {
            TextFormatApplication.engHeaderSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(7).getCell(1).toString());
        }
        if (sheet.getRow(8).getCell(1).toString().trim() != null) {
            TextFormatApplication.engNameTheme = sheet.getRow(8).getCell(1).toString();
        }
        if (sheet.getRow(9).getCell(1).toString().trim() != null) {
            TextFormatApplication.engNameSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(9).getCell(1).toString());
        }
        if (sheet.getRow(10).getCell(1).toString().trim() != null) {
            TextFormatApplication.engMajorTheme = sheet.getRow(10).getCell(1).toString();
        }
        if (sheet.getRow(11).getCell(1).toString().trim() != null) {
            TextFormatApplication.engMajorSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(11).getCell(1).toString());
        }
        if (sheet.getRow(12).getCell(1).toString().trim() != null) {
            TextFormatApplication.chnAbstractHeaderTheme = sheet.getRow(12).getCell(1).toString();
        }
        if (sheet.getRow(13).getCell(1).toString().trim() != null) {
            TextFormatApplication.chnAbstractTextTheme = sheet.getRow(13).getCell(1).toString();
        }
        if (sheet.getRow(14).getCell(1).toString().trim() != null) {
            TextFormatApplication.chnAbstractSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(14).getCell(1).toString());
        }
        if (sheet.getRow(15).getCell(1).toString().trim() != null) {
            TextFormatApplication.engAbstractHeaderTheme = sheet.getRow(15).getCell(1).toString();
        }
        if (sheet.getRow(16).getCell(1).toString().trim() != null) {
            TextFormatApplication.engAbstractTextTheme = sheet.getRow(16).getCell(1).toString();
        }
        if (sheet.getRow(17).getCell(1).toString().trim() != null) {
            TextFormatApplication.engAbstractSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(17).getCell(1).toString());
        }
        if (sheet.getRow(18).getCell(1).toString().trim() != null) {
            TextFormatApplication.chnKeywordsHeaderTheme = sheet.getRow(18).getCell(1).toString();
        }
        if (sheet.getRow(19).getCell(1).toString().trim() != null) {
            TextFormatApplication.chnKeywordsTextTheme = sheet.getRow(19).getCell(1).toString();
        }
        if (sheet.getRow(20).getCell(1).toString().trim() != null) {
            TextFormatApplication.chnKeywordsSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(20).getCell(1).toString());
        }
        if (sheet.getRow(21).getCell(1).toString().trim() != null) {
            TextFormatApplication.engKeywordsHeaderTheme = sheet.getRow(21).getCell(1).toString();
        }
        if (sheet.getRow(22).getCell(1).toString().trim() != null) {
            TextFormatApplication.engKeywordsTextTheme = sheet.getRow(22).getCell(1).toString();
        }
        if (sheet.getRow(23).getCell(1).toString().trim() != null) {
            TextFormatApplication.engKeywordsSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(23).getCell(1).toString());
        }
        if (sheet.getRow(24).getCell(1).toString().trim() != null) {
            TextFormatApplication.cataTheme = sheet.getRow(24).getCell(1).toString();
        }
        if (sheet.getRow(25).getCell(1).toString().trim() != null) {
            TextFormatApplication.cataSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(25).getCell(1).toString());
        }
        if (sheet.getRow(26).getCell(1).toString().trim() != null) {
            TextFormatApplication.cataTextTheme = sheet.getRow(26).getCell(1).toString();
        }
        if (sheet.getRow(27).getCell(1).toString().trim() != null) {
            TextFormatApplication.cataTextSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(27).getCell(1).toString());
        }
        if (sheet.getRow(28).getCell(1).toString().trim() != null) {
            TextFormatApplication.litTheme = sheet.getRow(28).getCell(1).toString();
        }
        if (sheet.getRow(29).getCell(1).toString().trim() != null) {
            TextFormatApplication.litSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(29).getCell(1).toString());
        }
        if (sheet.getRow(30).getCell(1).toString().trim() != null) {
            TextFormatApplication.chnlitTextTheme = sheet.getRow(30).getCell(1).toString();
        }
        if (sheet.getRow(31).getCell(1).toString().trim() != null) {
            TextFormatApplication.englitTextTheme = sheet.getRow(31).getCell(1).toString();
        }
        if (sheet.getRow(32).getCell(1).toString().trim() != null) {
            TextFormatApplication.litTextSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(32).getCell(1).toString());
        }
        if (sheet.getRow(33).getCell(1).toString().trim() != null) {
            TextFormatApplication.textHeader1Theme = sheet.getRow(33).getCell(1).toString();
        }
        if (sheet.getRow(34).getCell(1).toString().trim() != null) {
            TextFormatApplication.textHeader1Size = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(34).getCell(1).toString());
        }
        if (sheet.getRow(35).getCell(1).toString().trim() != null) {
            TextFormatApplication.textHeader2Theme = sheet.getRow(35).getCell(1).toString();
        }
        if (sheet.getRow(36).getCell(1).toString().trim() != null) {
            TextFormatApplication.textHeader2Size = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(36).getCell(1).toString());
        }
        if (sheet.getRow(37).getCell(1).toString().trim() != null) {
            TextFormatApplication.textTheme = sheet.getRow(37).getCell(1).toString();
        }
        if (sheet.getRow(38).getCell(1).toString().trim() != null) {
            TextFormatApplication.textSize = TextFormatApplication.wordFontSize
                    .get(sheet.getRow(38).getCell(1).toString());
        }
    }
}

package com.mng;

import com.mng.Bean.JsonBean;
import com.mng.Bean.VacateBean;
import com.mng.Dao.VacateDao;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletRequest;
import javax.sql.DataSource;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.sql.SQLException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("api")
public class VacateController {
    @Autowired
    DataSource dataSource;

    @Autowired
    VacateDao vacateDao;

    @RequestMapping("/vacate/insertrecord")
    public Map<String, Object> insertrecord(@RequestBody VacateBean vacateBean) throws SQLException {
        vacateDao.vacateAdd(vacateBean);
        Map<String, Object> tag = new HashMap<>();
        tag.put("result", "success");
        return tag;
    }

    @RequestMapping("/vacate/upload")
    public void uploadfile(@RequestParam("file") MultipartFile request) throws IOException {
        XWPFDocument document = new XWPFDocument(request.getInputStream());
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        if (paragraphs.size() != 0) {
            for (int i = 0; i < paragraphs.size(); i++) {
                System.out.println(paragraphs.get(i).getText());
            }
        }
    }

    @RequestMapping("/vacate/getvacatelist")
    public Map<String, Object> getvacatelist(@RequestBody JsonBean jsonBean) throws SQLException {
        Map<String, Object> tag = vacateDao.getVacateList(jsonBean.getUserid(), jsonBean.getName());
        return tag;
    }

    @RequestMapping("/vacate/updatevacate")
    public void updatevacate(@RequestBody VacateBean vacateBean) throws SQLException {
        vacateDao.updateVacate(vacateBean);
    }

    @RequestMapping("/vacate/changeagree")
    public void changeagree(@RequestBody VacateBean vacateBean) throws SQLException {
        vacateDao.change_agree(vacateBean);
    }

    @RequestMapping("/vacate/changerefuse")
    public void changerefuse(@RequestBody VacateBean vacateBean) throws SQLException {
        vacateDao.change_refuse(vacateBean);
    }
}

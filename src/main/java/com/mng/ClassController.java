package com.mng;

import com.mng.Bean.JsonBean;
import com.mng.Dao.ClassDao;
import com.mng.Dao.VacateDao;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.sql.DataSource;
import java.sql.SQLException;
import java.util.Map;

@RestController
@RequestMapping("api")
public class ClassController {
    @Autowired
    DataSource dataSource;

    @Autowired
    ClassDao classDao;

    @Autowired
    VacateDao vacateDao;

    @RequestMapping("getgrade")
    public Map<String, Object> getgrade() throws SQLException {
        Map<String, Object> tag = classDao.getGrade();
        return tag;
    }

    @RequestMapping("getclass")
    public Map<String, Object> getclass1() throws SQLException {
        Map<String, Object> tag = classDao.getClass2();
        return tag;
    }

    @RequestMapping("getvacate")
    public Map<String, Object> getvacate(@RequestBody JsonBean jsonBean) throws SQLException {
        Map<String, Object> tag = vacateDao.getVacate(jsonBean.getGrade(), jsonBean.getClass1());
        return tag;
    }
}

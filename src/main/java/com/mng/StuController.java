package com.mng;

import com.mng.Bean.JsonBean;
import com.mng.Bean.StuBean;
import com.mng.Dao.UserDao;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.sql.DataSource;
import java.sql.SQLException;
import java.util.Map;

@RestController
@RequestMapping("api")
public class StuController {
    @Autowired
    DataSource dataSource;

    @Autowired
    UserDao userDao;

    @RequestMapping("/stuinfo/queryinfo")
    public Map<String,Object> queryinfo(@RequestBody JsonBean jsonBean) throws SQLException{
        Map<String, Object> tag = userDao.currentUser_formUserName(jsonBean.getUserName());
        tag.putAll(userDao.getErrorList());
        return tag;
    }

    @RequestMapping("/stuinfo/upinfophone")
    public void upinfophone(@RequestBody StuBean stuBean) throws SQLException{
        userDao.infoUpdate_phone(stuBean.getUsername(),stuBean.getPhone(),stuBean.getBirthday(),stuBean.getRateScore());
    }

    @RequestMapping("/stuinfo/getinfolist")
    public Map<String,Object> getinfolist(@RequestBody JsonBean jsonBean) throws SQLException{
        Map<String, Object> tag = userDao.getInfoList(jsonBean.getGrade(), jsonBean.getClass1());
        return tag;
    }
}

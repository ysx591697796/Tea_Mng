package com.mng;

import com.mng.Bean.JsonBean;
import com.mng.Bean.VisiterBean;
import com.mng.Dao.UserDao;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.sql.DataSource;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping("api")
public class LoginController {
    private String username = null;

    @Autowired
    DataSource dataSource;

    @Autowired
    UserDao userDao;

    @RequestMapping("login")
    public Map<String, Object> login1(@RequestBody JsonBean adminBean) throws SQLException {
        Map<String, Object> tag = userDao.Login(adminBean.getUserName(), adminBean.getPassword(), adminBean.getType());
        if(tag.get("userName")!=null){
            username = (String)tag.get("userName");
        }
        return tag;
    }

    @RequestMapping("currentUser1")
    public Map<String, Object> currentUser1() throws SQLException {
        Map<String, Object> tag = userDao.currentUser(username);
        return tag;
    }
}

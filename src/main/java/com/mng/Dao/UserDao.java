package com.mng.Dao;

import org.omg.CORBA.StringSeqHelper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.relational.core.sql.SQL;
import org.springframework.web.bind.annotation.RestController;

import javax.sql.DataSource;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

@RestController
public class UserDao {
    @Autowired
    DataSource dataSource;

    public Map<String, Object> Login(String username, String password, String type) throws SQLException {
        Map<String, Object> tag = new HashMap<>();
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        ResultSet rs = null;
        String sql = "select * from stuinfo where username='" + username + "' and password='" + password + "';";
        stm = connection.prepareStatement(sql);
        rs = stm.executeQuery();
        if (rs.next()) {
            if (rs.getString("status").equals("1")) {
                tag.put("currentAuthority", "admin");
            } else {
                tag.put("currentAuthority", "user");
            }
            tag.put("status", "ok");
            tag.put("type", type);
            tag.put("userName", username);
        } else {
            tag.put("currentAuthority", "guest");
            tag.put("status", "error");
            tag.put("type", type);
        }
        rs.close();
        stm.close();
        connection.close();
        return tag;
    }

    public Map<String, Object> currentUser(String username) throws SQLException {
        Map<String, Object> tag = new HashMap<>();
        if (username != null) {
            Connection connection = dataSource.getConnection();
            PreparedStatement stm = null;
            ResultSet rs = null;
            String sql = "select * from stuinfo where username='" + username + "';";
            stm = connection.prepareStatement(sql);
            rs = stm.executeQuery();
            if (rs.next()) {
                tag.put("avatar", "https://gw.alipayobjects.com/zos/antfincdn/XAosXuNZyF/BiazfanxmamNRoxxVxka.png");
                tag.put("username", username);
                tag.put("name", rs.getString("name"));
                tag.put("userid", rs.getString("id"));
                tag.put("sex", rs.getString("sex"));
                tag.put("class1", rs.getString("class"));
                tag.put("politics", rs.getString("politics"));
                tag.put("phone", rs.getString("phone"));
                tag.put("birthday", rs.getString("birthday"));
                tag.put("grade", rs.getString("grade"));
            }
            rs.close();
            stm.close();
            connection.close();
        } else {
            tag = null;
        }
        return tag;
    }

    public Map<String, Object> currentUser_formName(String name) throws SQLException {
        Map<String, Object> tag = new HashMap<>();
        if (name != null) {
            Connection connection = dataSource.getConnection();
            PreparedStatement stm = null;
            ResultSet rs = null;
            String sql = "select * from stuinfo where name='" + name + "';";
            stm = connection.prepareStatement(sql);
            rs = stm.executeQuery();
            if (rs.next()) {
                tag.put("username", rs.getString("username"));
                tag.put("name", name);
                tag.put("userid", rs.getString("id"));
                tag.put("sex", rs.getString("sex"));
                tag.put("class1", rs.getString("class"));
                tag.put("politics", rs.getString("politics"));
                tag.put("phone", rs.getString("phone"));
                tag.put("birthday", rs.getString("birthday"));
                tag.put("grade", rs.getString("grade"));
            }
            rs.close();
            stm.close();
            connection.close();
        } else {
            tag = null;
        }
        return tag;
    }

    public Map<String, Object> currentUser_formUserName(String username) throws SQLException {
        Map<String, Object> tag = new HashMap<>();
        if (username != null) {
            Connection connection = dataSource.getConnection();
            PreparedStatement stm = null;
            ResultSet rs = null;
            String sql = "select * from stuinfo where username='" + username + "';";
            stm = connection.prepareStatement(sql);
            rs = stm.executeQuery();
            if (rs.next()) {
                tag.put("username", username);
                tag.put("name", rs.getString("name"));
                tag.put("userid", rs.getString("id"));
                tag.put("sex", rs.getString("sex"));
                tag.put("class1", rs.getString("class"));
                tag.put("politics", rs.getString("politics"));
                tag.put("phone", rs.getString("phone"));
                tag.put("birthday", rs.getString("birthday"));
                tag.put("grade", rs.getString("grade"));
                tag.put("avatar", "https://gw.alipayobjects.com/zos/antfincdn/XAosXuNZyF/BiazfanxmamNRoxxVxka.png");
                tag.put("ratescore", rs.getString("rate_score"));
                tag.put("ratecount", rs.getString("rate_count"));
                tag.put("testcount", rs.getString("test_count"));
            }
            rs.close();
            stm.close();
            connection.close();
        } else {
            tag = null;
        }
        return tag;
    }

    public void infoUpdate(String username, String email, String phone, String birthday, String address) throws SQLException {
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        String sql = "update stuinfo set email=?,phone=?,birthday=?,address=? where username=?;";
        stm = connection.prepareStatement(sql);
        stm.setString(1, email);
        stm.setString(2, phone);
        stm.setString(3, birthday);
        stm.setString(4, address);
        stm.setString(5, username);
        stm.executeUpdate();
        stm.close();
        connection.close();
    }

    public void infoUpdate_phone(String username, String phone, String birthday, Float score) throws SQLException {
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        String sql;
        if (score != null) {
            sql = "update stuinfo set rate_score=rate_score+?,rate_count=rate_count+1 where username=?;";
            stm = connection.prepareStatement(sql);
            stm.setString(1, String.valueOf(score));
            stm.setString(2, username);
        } else {
            if (phone != null) {
                sql = "update stuinfo set phone=?,birthday=? where username=?;";
                stm = connection.prepareStatement(sql);
                stm.setString(1, phone);
                stm.setString(2, birthday);
                stm.setString(3, username);
            }else {
                sql = "update stuinfo set test_count=test_count+1 where username=?;";
                stm = connection.prepareStatement(sql);
                stm.setString(1, username);
            }
        }
        stm.execute();
        stm.close();
        connection.close();
    }

    public Map<String, Object> getInfoList(String grade, String class1) throws SQLException {
        Map<String, Object> tag = new HashMap<>();
        ArrayList<Object> list = new ArrayList<Object>();
        Map<String, Object> info = null;
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        ResultSet rs = null;
        String sql = "select * from stuinfo where grade = '" + grade + "' and class = '" + class1 + "';";
        stm = connection.prepareStatement(sql);
        rs = stm.executeQuery();
        while (rs.next()) {
            info = new HashMap<>();
            info.put("username", rs.getString("username"));
            info.put("name", rs.getString("name"));
            info.put("userid", rs.getString("id"));
            info.put("sex", rs.getString("sex"));
            info.put("class1", rs.getString("class"));
            info.put("politics", rs.getString("politics"));
            info.put("phone", rs.getString("phone"));
            info.put("birthday", rs.getString("birthday"));
            info.put("grade", rs.getString("grade"));
//            info.put("avatar","https://gw.alipayobjects.com/zos/antfincdn/XAosXuNZyF/BiazfanxmamNRoxxVxka.png");
            list.add(info);
        }
        tag.put("infoList1", list);
        rs.close();
        stm.close();
        connection.close();
        return tag;
    }
}

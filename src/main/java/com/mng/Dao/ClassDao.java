package com.mng.Dao;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RestController;

import javax.sql.DataSource;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
public class ClassDao {
    @Autowired
    DataSource dataSource;

    public Map<String, Object> getGrade() throws SQLException {
        Map<String, Object> tag = new HashMap<>();
        ArrayList<String> str = new ArrayList<String>();
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        ResultSet rs = null;
        String sql = "select * from grade;";
        stm = connection.prepareStatement(sql);
        rs = stm.executeQuery();
        while (rs.next()) {
            str.add(rs.getString("grade"));
        }
        tag.put("gradeList1", str);
        rs.close();
        stm.close();
        connection.close();
        return tag;
    }

    public Map<String, Object> getClass2() throws SQLException {
        Map<String, Object> tag = new HashMap<>();
        ArrayList<String> str = new ArrayList<String>();
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        ResultSet rs = null;
        String sql = "select name from class ;";
        stm = connection.prepareStatement(sql);
        rs = stm.executeQuery();
        while (rs.next()) {
            str.add(rs.getString("name"));
        }
        tag.put("name", str);
        rs.close();
        stm.close();
        connection.close();
        return tag;
    }

    public Map<String, Object> getClass1(String grade) throws SQLException {
        Map<String, Object> tag = new HashMap<>();
        ArrayList<String> str = new ArrayList<String>();
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        ResultSet rs = null;
        String sql = "select * from class where grade = '" + grade + "';";
        stm = connection.prepareStatement(sql);
        rs = stm.executeQuery();
        while (rs.next()) {
            str.add(rs.getString("class"));
        }
        tag.put("class", str);
        rs.close();
        stm.close();
        connection.close();
        return tag;
    }

    public Map<String, Object> getClass1Stu(String grade, String class1) throws SQLException {
        Map<String, Object> tag = new HashMap<>();
        ArrayList<String> str = new ArrayList<String>();
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        ResultSet rs = null;
        String sql = "select * from stuinfo where grade = '" + grade + "' and class = '" + class1 + "';";
        stm = connection.prepareStatement(sql);
        rs = stm.executeQuery();
        while (rs.next()) {
            str.add(rs.getString("name"));
        }
        tag.put("class_stu", str);
        rs.close();
        stm.close();
        connection.close();
        return tag;
    }
}

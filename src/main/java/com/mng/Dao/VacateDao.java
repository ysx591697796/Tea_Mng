package com.mng.Dao;

import com.mng.Bean.VacateBean;
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
public class VacateDao {
    @Autowired
    DataSource dataSource;

    public Map<String, Object> getVacate(String grade, String class1) throws SQLException {
        Map<String, Object> tag = new HashMap<>();
        ArrayList<Object> str = new ArrayList<Object>();
        Map<String, Object> vacate = null;
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        ResultSet rs = null;
        String sql = "select * from vacaterecord where grade = '" + grade + "' and class = '" + class1 + "';";
        stm = connection.prepareStatement(sql);
        rs = stm.executeQuery();
        while (rs.next()) {
            vacate = new HashMap<>();
            vacate.put("name", rs.getString("name"));
            vacate.put("id", rs.getString("id"));
            vacate.put("class1", rs.getString("class"));
            vacate.put("grade", rs.getString("grade"));
            vacate.put("reason", rs.getString("reason"));
            vacate.put("starttime", rs.getString("starttime"));
            vacate.put("endtime", rs.getString("endtime"));
            vacate.put("phone", rs.getString("phone"));
            vacate.put("status", rs.getString("status"));
            vacate.put("type", rs.getString("type"));
            vacate.put("result", rs.getString("result"));
            vacate.put("vid", rs.getString("vid"));
            str.add(vacate);
        }
        tag.put("vacate", str);
        rs.close();
        stm.close();
        connection.close();
        return tag;
    }

    public Map<String, Object> getVacateInfo(String vid) throws SQLException {
        Map<String, Object> tag = new HashMap<>();
        ArrayList<String> str = new ArrayList<String>();
//        Map<String, Object> vacate = new HashMap<>();
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        ResultSet rs = null;
        String sql = "select * from vacaterecord where vid = '" + vid + "';";
        stm = connection.prepareStatement(sql);
        rs = stm.executeQuery();
        if (rs.next()) {
//            vacate.put("name",rs.getString("name"));
//            vacate.put("id",rs.getString("id"));
//            vacate.put("phone",rs.getString("phone"));
//            vacate.put("status",rs.getString("status"));
//            vacate.put("vid",rs.getString("vid"));
            str.add(rs.getString("name"));
            str.add(rs.getString("id"));
            str.add(rs.getString("class"));
            str.add(rs.getString("phone"));
            str.add(rs.getString("grade"));
            str.add(rs.getString("reason"));
            str.add(rs.getString("starttime"));
            str.add(rs.getString("endtime"));
            str.add(rs.getString("type"));
            str.add(rs.getString("status"));
            str.add(rs.getString("result"));
        }
        tag.put("vacate", str);
        rs.close();
        stm.close();
        connection.close();
        return tag;
    }

    public Map<String, Object> getVacateInfo_fromStu(String name) throws SQLException {
        Map<String, Object> tag = new HashMap<>();
        ArrayList<String> str = new ArrayList<String>();
//        Map<String, Object> vacate = new HashMap<>();
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        ResultSet rs = null;
        String sql = "select * from vacaterecord where name = '" + name + "';";
        stm = connection.prepareStatement(sql);
        rs = stm.executeQuery();
        if (rs.next()) {
//            vacate.put("name",rs.getString("name"));
//            vacate.put("id",rs.getString("id"));
//            vacate.put("phone",rs.getString("phone"));
//            vacate.put("status",rs.getString("status"));
//            vacate.put("vid",rs.getString("vid"));
            str.add(rs.getString("name"));
            str.add(rs.getString("id"));
            str.add(rs.getString("class"));
            str.add(rs.getString("phone"));
            str.add(rs.getString("grade"));
            str.add(rs.getString("reason"));
            str.add(rs.getString("starttime"));
            str.add(rs.getString("endtime"));
            str.add(rs.getString("type"));
            str.add(rs.getString("status"));
            str.add(rs.getString("result"));
        }
        tag.put("vacate", str);
        rs.close();
        stm.close();
        connection.close();
        return tag;
    }

    public void vacateAdd(VacateBean vacateBean) throws SQLException {
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        String sql = "insert into vacaterecord(name,id,class,phone,grade," +
                "reason,starttime,endtime,type) values(?,?,?,?,?,?,?,?,?);";
        stm = connection.prepareStatement(sql);
        stm.setString(1,vacateBean.getName());
        stm.setString(2,vacateBean.getId());
        stm.setString(3,vacateBean.getClass1());
        stm.setString(4,vacateBean.getPhone());
        stm.setString(5,vacateBean.getGrade());
        stm.setString(6,vacateBean.getReason());
        stm.setString(7,vacateBean.getStarttime());
        stm.setString(8,vacateBean.getEndtime());
        stm.setString(9,vacateBean.getType());
        stm.execute();
        stm.close();
        connection.close();
    }

    public Map<String,Object> getVacateList(String id,String name) throws SQLException{
        Map<String, Object> tag = new HashMap<>();
        ArrayList<Object> str = new ArrayList<Object>();
        Map<String, Object> vacate = null;
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        ResultSet rs = null;
        String sql = "select * from vacaterecord where id = '" + id + "' and name = '" + name + "';";
        stm = connection.prepareStatement(sql);
        rs = stm.executeQuery();
        while (rs.next()) {
            vacate = new HashMap<>();
            vacate.put("name", rs.getString("name"));
            vacate.put("id", rs.getString("id"));
            vacate.put("class1", rs.getString("class"));
            vacate.put("grade", rs.getString("grade"));
            vacate.put("reason", rs.getString("reason"));
            vacate.put("starttime", rs.getString("starttime"));
            vacate.put("endtime", rs.getString("endtime"));
            vacate.put("phone", rs.getString("phone"));
            vacate.put("status", rs.getString("status"));
            vacate.put("type", rs.getString("type"));
            vacate.put("result", rs.getString("result"));
            vacate.put("vid", rs.getString("vid"));
            str.add(vacate);
        }
        tag.put("vacate", str);
        rs.close();
        stm.close();
        connection.close();
        return tag;
    }

    public void updateVacate(VacateBean vacateBean) throws SQLException{
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        String sql = "update vacaterecord set phone=?,reason=?,starttime=?,endtime=?,type=?,result=?" +
                " where vid = '" + vacateBean.getVid() + "' and name = '" + vacateBean.getName() + "';";
        stm = connection.prepareStatement(sql);
        stm.setString(1, vacateBean.getPhone());
        stm.setString(2, vacateBean.getReason());
        stm.setString(3, vacateBean.getStarttime());
        stm.setString(4, vacateBean.getEndtime());
        stm.setString(5, vacateBean.getType());
        stm.setString(6, "未审核");
        stm.execute();
        stm.close();
        connection.close();
    }

    public void change_agree(VacateBean vacateBean) throws SQLException{
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        String sql = "update vacaterecord set result=? where vid = '" + vacateBean.getVid() + "';";
        stm = connection.prepareStatement(sql);
        stm.setString(1, "同意");
        stm.execute();
        stm.close();
        connection.close();
    }

    public void change_refuse(VacateBean vacateBean) throws SQLException{
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        String sql = "update vacaterecord set result=? where vid = '" + vacateBean.getVid() + "';";
        stm = connection.prepareStatement(sql);
        stm.setString(1, "驳回");
        stm.execute();
        stm.close();
        connection.close();
    }
}

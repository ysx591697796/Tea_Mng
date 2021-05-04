package com.mng.Dao;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RestController;

import javax.sql.DataSource;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;

@RestController
public class ErrorDao {
    @Autowired
    DataSource dataSource;

    public void errorIncrease(String str) throws SQLException {
        Connection connection = dataSource.getConnection();
        PreparedStatement stm = null;
        String sql;
        sql = "update error_num set " + str + "=" + str + "+1 where id=1;";
        stm = connection.prepareStatement(sql);
        stm.execute();
        stm.close();
        connection.close();
    }
}

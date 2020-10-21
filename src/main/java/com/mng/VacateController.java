package com.mng;

import com.mng.Bean.JsonBean;
import com.mng.Bean.VacateBean;
import com.mng.Dao.VacateDao;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.sql.DataSource;
import java.sql.SQLException;
import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping("api")
public class VacateController {
    @Autowired
    DataSource dataSource;

    @Autowired
    VacateDao vacateDao;

    @RequestMapping("/vacate/insertrecord")
    public Map<String,Object> insertrecord(@RequestBody VacateBean vacateBean) throws SQLException {
        vacateDao.vacateAdd(vacateBean);
        Map<String, Object> tag = new HashMap<>();
        tag.put("result","success");
        return tag;
    }

    @RequestMapping("/vacate/getvacatelist")
    public Map<String,Object> getvacatelist(@RequestBody JsonBean jsonBean) throws SQLException{
        Map<String, Object> tag = vacateDao.getVacateList(jsonBean.getUserid(),jsonBean.getName());
        return tag;
    }

    @RequestMapping("/vacate/updatevacate")
    public void updatevacate(@RequestBody VacateBean vacateBean) throws SQLException{
        vacateDao.updateVacate(vacateBean);
    }

    @RequestMapping("/vacate/changeagree")
    public void changeagree(@RequestBody VacateBean vacateBean) throws SQLException{
        vacateDao.change_agree(vacateBean);
    }

    @RequestMapping("/vacate/changerefuse")
    public void changerefuse(@RequestBody VacateBean vacateBean) throws SQLException{
        vacateDao.change_refuse(vacateBean);
    }
}

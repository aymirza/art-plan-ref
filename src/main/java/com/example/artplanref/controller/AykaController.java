package com.example.artplanref.controller;

import com.example.artplanref.model.OylikPlan;
import com.example.artplanref.model.entity.OylikPlanEntity;
import com.example.artplanref.repository.OylikPlanRepository;
import com.example.artplanref.service.OylikPlanService;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.sql.Date;
import java.util.ArrayList;
import java.util.List;

@RestController
public class AykaController {

    @Autowired
    private OylikPlanRepository oylikPlanRepository;

    @Autowired
    private OylikPlanService oylikPlanService;


    @GetMapping("/hello")
    public String index() {
        return "hello";
    }


    @PostMapping("/import-oylik-plan")
    public List<OylikPlan> importOylikPlan(@RequestParam("file") MultipartFile files) throws IOException {
        List<OylikPlan> oylikPlans = new ArrayList<>();
        XSSFWorkbook workbook = new XSSFWorkbook((files.getInputStream()));
        // Read oylik plan data form excel sheet1;
        XSSFSheet worksheet = workbook.getSheetAt(0);
        for (int index = 0; index < worksheet.getPhysicalNumberOfRows(); index++) {
            if (index > 0) {
                XSSFRow row = worksheet.getRow(index);
                OylikPlan oylikPlan = new OylikPlan();
                oylikPlan.planDate = Date.valueOf(getCellValue(row, 0));
                oylikPlan.sapCode = getCellValue(row, 1);
                oylikPlan.fullNameModel = getCellValue(row, 2);
                oylikPlan.nameModel = getCellValue(row, 3);
                oylikPlan.brand = getCellValue(row, 4);
                oylikPlan.color = getCellValue(row, 5);
                oylikPlan.shipment = getCellValue(row, 6);
                oylikPlan.quantity = Integer.valueOf(getCellValue(row, 7));
                oylikPlans.add(oylikPlan);
            }
        }


        // Save to db
        List<OylikPlanEntity> oylikPlanEntities = new ArrayList<>();
        if (oylikPlans.size() > 0){
            oylikPlans.forEach(x->{
                OylikPlanEntity oylikPlanEntity = new OylikPlanEntity();
                oylikPlanEntity.plan_date = x.planDate;
                oylikPlanEntity.sap_code = x.sapCode;
                oylikPlanEntity.full_name_model = x.fullNameModel;
                oylikPlanEntity.name_model = x.nameModel;
                oylikPlanEntity.brand = x.brand;
                oylikPlanEntity.color = x.color;
                oylikPlanEntity.shipment = x.shipment;
                oylikPlanEntity.quantity = x.quantity;
                oylikPlanEntities.add(oylikPlanEntity);
            });
            oylikPlanRepository.saveAll(oylikPlanEntities);
        }
        return oylikPlans;
    }


    private String getCellValue(Row row, int cellNo){
        DataFormatter formatter = new DataFormatter();
        Cell cell = row.getCell(cellNo);
        return formatter.formatCellValue(cell);
    }



}

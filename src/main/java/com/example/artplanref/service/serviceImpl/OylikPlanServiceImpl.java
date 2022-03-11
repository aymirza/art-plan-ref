package com.example.artplanref.service.serviceImpl;

import com.example.artplanref.model.entity.OylikPlanEntity;
import com.example.artplanref.repository.OylikPlanRepository;
import com.example.artplanref.service.OylikPlanService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class OylikPlanServiceImpl implements OylikPlanService {

    @Autowired
    private OylikPlanRepository oylikPlanRepository;


    @Override
    public List<OylikPlanEntity> getAllOylikPlan() {
        return oylikPlanRepository.findAll();
    }
}

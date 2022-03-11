package com.example.artplanref.repository;

import com.example.artplanref.model.entity.OylikPlanEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface OylikPlanRepository extends JpaRepository<OylikPlanEntity, Long> {
}

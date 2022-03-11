package com.example.artplanref.service.serviceImpl;

import com.example.artplanref.model.entity.StudentEntity;
import com.example.artplanref.repository.StudentRepository;
import com.example.artplanref.service.StudentService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class StudentServiceImpl implements StudentService {

    @Autowired
    private StudentRepository studentRepository;


    @Override
    public List<StudentEntity> getAllStudents() {
        return studentRepository.findAll();
    }
}

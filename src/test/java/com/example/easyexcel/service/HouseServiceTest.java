package com.example.easyexcel.service;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import static org.junit.jupiter.api.Assertions.*;

@SpringBootTest
class HouseServiceTest {

    @Autowired
    private HouseService houseService;

    @Test
    void export2Excel() {
        houseService.export2Excel();
    }
}
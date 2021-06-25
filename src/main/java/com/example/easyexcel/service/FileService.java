package com.example.easyexcel.service;

import org.springframework.stereotype.Service;

import java.io.File;

/**
 * @author 92339
 */
@Service
public class FileService {

    private static final String OSS_URL = "https://minio.xx.com/";

    public String upload(File file, String fileName){
        return OSS_URL + fileName;
    }

}

package com.example.demo.Controller;
import com.example.demo.Service.Exercise1Service;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@RestController
public class Exercise1Controller {

    @Autowired
    Exercise1Service exercise1Service;

        private static final String UPLOAD_DIR = "C:\\Users\\Admin\\Desktop\\Demo.xlsx";

        @PostMapping("/uploads")
        @ResponseBody
        public String handleFileUpload() {
            try{
                exercise1Service.processExcelFile(UPLOAD_DIR);// sorting file column
                return "Process completed.";
            } catch (IOException e) {
                return "Process failed .... " + e.getMessage();
            }
        }

}

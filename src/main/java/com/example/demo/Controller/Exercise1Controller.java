package com.example.demo.Controller;
import com.example.demo.Service.Exercise1Service;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@RestController
public class Exercise1Controller {

    @Autowired
    Exercise1Service exercise1Service;

        @PostMapping("/upload/path")
        public String handleFileUpload(@RequestParam("filePath") String filePath) {
            try{
                exercise1Service.processExcelFile(filePath);
                return "Process completed.";
            } catch (IOException e) {
                return "Process failed .... " + e.getMessage();
            }
        }

    @PostMapping("/upload/file")
    public String handleFileUploadMultiPart(@RequestParam("file") MultipartFile file) {
        try{
            exercise1Service.processExcelFileMultiPart(file);
            return "Process completed.";
        } catch (IOException e) {
            return "Process failed .... " + e.getMessage();
        }
    }

}

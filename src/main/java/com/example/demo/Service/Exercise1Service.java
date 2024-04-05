package com.example.demo.Service;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

@Service
public class Exercise1Service {

    private final Logger logger = Logger.getLogger(Exercise1Service.class.getName());

    public void processExcelFile(String file) throws IOException {
            try  {
                logger.info(":: File uploading started ::");
                FileInputStream fileInputStream = new FileInputStream(new File(file));
                logger.info(":: File "+new File(file).getName()+" uploaded successfully and started reading it :: ");
                Workbook workbook = new XSSFWorkbook(fileInputStream);
                logger.info(":: Reading file sheet ::");
                Sheet sheet = workbook.getSheetAt(0); // Assuming first sheet

                logger.info(sheet.getWorkbook()+" "+sheet.getSheetName());
                // Read data from the first column
                TreeMap<Double,String> columnData = readFirstColumn(sheet);

                logger.info(":: Reading file completed ::");

                // Sort data in descending order
                logger.info(":: sorted data :: "+columnData.toString());

                logger.info(":: Writing file invoked ::");
                // Write sorted data back to the first column
                writeFirstColumn(sheet, columnData);

                logger.info(":: Writing file completed ::");

                FileOutputStream updatedfile = new FileOutputStream(file);
                workbook.write(updatedfile);
                logger.info(":: Data Copied to Excel ::");
                fileInputStream.close();
                updatedfile.close();
                fileRename(new File(file));

            }
            catch(Exception exception){
                logger.log(Level.SEVERE,exception.getMessage());
            }
        }

    public void processExcelFileMultiPart(MultipartFile file) throws IOException {
        try  {
                logger.info(":: File uploading started ::");
                FileInputStream fileInputStream = new FileInputStream(new File(Objects.requireNonNull(file.getOriginalFilename())).getAbsolutePath());
                logger.info(":: File " + file.getName() + " uploaded successfully and started reading it :: ");
                Workbook workbook = new XSSFWorkbook(fileInputStream);
                logger.info(":: Reading file sheet ::");
                Sheet sheet = workbook.getSheetAt(0); // Assuming first sheet

                logger.info(sheet.getWorkbook() + " " + sheet.getSheetName());
                // Read data from the first column
                TreeMap<Double, String> columnData = readFirstColumn(sheet);

                logger.info(":: Reading file completed ::");

                // Sort data in descending order
                logger.info(":: sorted data :: " + columnData.toString());

                logger.info(":: Writing file invoked ::");
                // Write sorted data back to the first column
                writeFirstColumn(sheet, columnData);

                logger.info(":: Writing file completed ::");

                FileOutputStream updatedfile = new FileOutputStream(new File(file.getOriginalFilename()).getAbsolutePath());
                workbook.write(updatedfile);
                logger.info(":: Data Copied to Excel ::");
                fileInputStream.close();
                updatedfile.close();
                fileRename(new File(file.getOriginalFilename()));


        }
        catch(Exception exception){
            logger.log(Level.SEVERE,exception.getMessage());
        }
    }

        private TreeMap<Double,String> readFirstColumn(Sheet sheet) {
            TreeMap<Double,String> columnData = new TreeMap<>();
//            int pointer = 0;
//                logger.info(String.valueOf(sheet.getActiveCell().getRow()+" "+sheet.getActiveCell().getColumn()));
            for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                Double key = (Double) sheet.getRow(r)
                        .getCell(0)
                        .getNumericCellValue();
                String value = sheet.getRow(r)
                        .getCell(1)
                        .getStringCellValue();
                columnData.put(key, value);
            }
            return columnData;
        }

        private void writeFirstColumn(Sheet sheet,  TreeMap<Double,String>  data) throws FileNotFoundException {
            int rowno=0;

            for(HashMap.Entry entry:data.entrySet()) {
                XSSFRow row= (XSSFRow) sheet.createRow(rowno++);
                row.createCell(0).setCellValue((Double)entry.getKey());
                row.createCell(1).setCellValue((String)entry.getValue());
            }

        }


        private void fileRename(File file){
            logger.info(":: File Rename Invoked ::");
            Path sourcePath = Paths.get(file.getPath());
            if(Files.exists(sourcePath)) {
                LocalDateTime currentDateTime = LocalDateTime.now();
                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss");
                String formattedDateTime = currentDateTime.format(formatter);
                String newFileName = formattedDateTime + "_" + file.getName();
                Path destinationPath = sourcePath.resolveSibling(newFileName);
//                File newFile = new File(file.getParent(), newFileName);
                try {
                    Files.move(sourcePath,destinationPath);
                    logger.info(":: File Renamed successfully to " + formattedDateTime + "_" + file.getName() + " ::");
                }
                catch(Exception e){
                    logger.warning(":: Internal Error :: "+e.getMessage());
                }
            }
            else{
                logger.log(Level.WARNING,":: Path not resolved ::");
            }
        }

    }


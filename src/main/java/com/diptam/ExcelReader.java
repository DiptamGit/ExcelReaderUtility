package com.diptam;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Optional;

public class ExcelReader {

    private static final String MASTER_FILE_PATH = "C:/Users/DiptamSarkar/Documents/SF/Performer/Master.xlsx";
    private static final String SCRENNER_FILE_PATH = "C:/Users/DiptamSarkar/Documents/SF/Performer/Screener.xlsx";
    private static final String FINAL_FILE_PATH = "C:/Users/DiptamSarkar/Documents/SF/Performer/Master_Screener.txt";

    public static void main(String[] args) {
        try {
            ArrayList<Master> masters = readMaster();
            ArrayList<Performer> performers = readScreener(masters);
            writeToFile(performers);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void writeToFile(ArrayList<Performer> performers) throws IOException {
        System.out.println("Writing to file started->");
        Path path = Paths.get(FINAL_FILE_PATH);
        try(BufferedWriter writer = Files.newBufferedWriter(path)) {
            performers.forEach(performer -> {
                try {
                    writer.write(performer.getFolio_id()+"|"+
                                 performer.getAssignee()+"|"+
                                 performer.getType()+"|"+
                                 performer.getAlias()+"|"+
                                 performer.getRole()+"|"+
                                 performer.getManager()+"|"+
                                 performer.getManager_alias()+"|"+
                                 performer.getLastupdated()+"|"+
                                 performer.getComment());
                    writer.newLine();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            });
        }
        System.out.println("Writing to file completed ->");
    }

    private static ArrayList<Performer> readScreener(ArrayList<Master> masters) throws Exception{
        System.out.println("Reading perfomer started->");
        DateTimeFormatter format = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:SS.ss");
        String referenceDate = "2018-03-28 00:00:00.00";
        Workbook workbook = WorkbookFactory.create(new File(SCRENNER_FILE_PATH));
        ArrayList<Performer> performers = new ArrayList<>();
        DataFormatter formatter = new DataFormatter();
        workbook.sheetIterator().forEachRemaining(sheet -> {
            sheet.forEach(row -> {
                Optional<Master> matchedMaster = masters.stream().filter(master -> master.getFolio_id().equals(formatter.formatCellValue(row.getCell(0)))
                                                                && master.getAssignee().equalsIgnoreCase(formatter.formatCellValue(row.getCell(3)))).findFirst();
                if (matchedMaster.isPresent()){
                    LocalDateTime lastUpdateAtMaster = LocalDateTime.parse(matchedMaster.get().getLastUpdated(), format);
                    LocalDateTime lastupdatedAtScreener = LocalDateTime.parse(formatter.formatCellValue(row.getCell(7)), format);
                    if(lastUpdateAtMaster.isAfter(lastupdatedAtScreener)){
                         performers.add(createPerformer(row, lastUpdateAtMaster, "Folio id found in master data with updated date"));
                    }else{
                        performers.add(createPerformer(row, lastupdatedAtScreener, "No Change Needed"));
                    }
                }else {
                    LocalDateTime lastupdatedAtScreener = LocalDateTime.parse(formatter.formatCellValue(row.getCell(7)), format);
                    LocalDateTime referenceDateTime = LocalDateTime.parse(referenceDate, format);
                    if(referenceDateTime.isAfter(lastupdatedAtScreener)){
                        performers.add(createPerformer(row, lastupdatedAtScreener, "last updated date is found before 2018-03-28"));
                    }else {
                        performers.add(createPerformer(row, lastupdatedAtScreener, "Bad Data needs further checking"));
                    }
                }
            });
        });
        System.out.println("Reading performer completed->");
        System.out.println("Final Performer List size : "+performers.size());
        workbook.close();
        return performers;
    }

    private static Performer createPerformer(Row row, LocalDateTime lastUpdated, String comment) {
        DateTimeFormatter format = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:SS.ss");
        //System.out.println(lastUpdated.format(format));
        DataFormatter formatter = new DataFormatter();
        Performer performer = new Performer();
        performer.setFolio_id(formatter.formatCellValue(row.getCell(0)));
        performer.setAssignee(formatter.formatCellValue(row.getCell(1)));
        performer.setType(formatter.formatCellValue(row.getCell(2)));
        performer.setAlias(formatter.formatCellValue(row.getCell(3)));
        performer.setRole(formatter.formatCellValue(row.getCell(4)));
        performer.setManager(formatter.formatCellValue(row.getCell(5)));
        performer.setManager_alias(formatter.formatCellValue(row.getCell(6)));
        performer.setLastupdated(lastUpdated.format(format));
        performer.setComment(comment);
        //System.out.println(performer);

        return performer;
    }

    private static ArrayList<Master> readMaster() throws Exception{
        System.out.println("Reading Master ->");
        Workbook workbook = WorkbookFactory.create(new File(MASTER_FILE_PATH));
        ArrayList<Master> masters = new ArrayList<>();
        DataFormatter formatter = new DataFormatter();
        workbook.sheetIterator().forEachRemaining(sheet -> {
            sheet.forEach(row -> {
                Master master = new Master();
                master.setFolio_id(formatter.formatCellValue(row.getCell(0)));
                master.setLastUpdated(formatter.formatCellValue(row.getCell(1)));
                master.setAssignee(formatter.formatCellValue(row.getCell(2)));
                masters.add(master);
                //System.out.println(master);
            });
        });
        workbook.close();
        System.out.println("Reading Master completed->");
        System.out.println("Master record size : "+masters.size());
        return masters;
    }
}

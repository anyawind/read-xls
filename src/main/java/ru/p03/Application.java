package ru.p03;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Date;

import static java.util.Objects.isNull;

public class Application {
    private static final Logger logger = LoggerFactory.getLogger(Application.class);
    private static ApplicationProperties properties = ApplicationProperties.getInstance();
    private static String url = properties.getProperty("url");
    private static String user = properties.getProperty("user");
    private static String password = properties.getProperty("password");
    private static String way = properties.getProperty("way");

    public static void main( String[] args ) {
        Connection connection = null;
        try {
            XSSFWorkbook xlsx = new XSSFWorkbook();
            XSSFSheet sheet = xlsx.createSheet();
            XSSFRow row;
            row = sheet.createRow(0);
            row.createCell(0).setCellValue("снилс");
            row.createCell(1).setCellValue("id");
            row.createCell(2).setCellValue("фамилия");
            row.createCell(3).setCellValue("имя");
            row.createCell(4).setCellValue("отчество");
            row.createCell(6).setCellValue("дата рождения");
            row.createCell(7).setCellValue("категория");

            Class.forName("com.ibm.db2.jcc.DB2Driver");
            connection = DriverManager.getConnection(url, user, password);
            Statement stat = connection.createStatement();

            Date date = new Date();

            FileInputStream vyp = new FileInputStream("C:/test/mvd08/Копия Список на 01.01.2022г.xls");
            Workbook wb = new HSSFWorkbook(vyp);
            int n = wb.getSheetAt(0).getLastRowNum();

            SimpleDateFormat sqa = new SimpleDateFormat("dd.MM.yyyy");
            SimpleDateFormat sqa1 = new SimpleDateFormat("yyyy-MM-dd");

            logger.info("получение данных " + date);
            int j = 0, f = 1;
           for (int i = 0; i < n; i++) {
               int pe = 0;
               row = sheet.createRow(f);
                   if(!isNull(wb.getSheetAt(0).getRow(i+1).getCell(1))) {
                       System.out.println(wb.getSheetAt(0).getRow(i + 1).getCell(1).getStringCellValue());
                       String query = String.format("SELECT t1.id, t1.fa, t1.im, t1.ot, t1.npers, t1.rdat, t2.kat \n" +
                               "\tFROM table1 t1\n" +
                               "left join table2 t2 on t1.id=t2.id.id\n" +
                               "where t1.npers='%s' and p.dat=(select max(dat) from table2 where id=t1.id) and t1.vibr=1\n" +
                               "fetch first 1 rows only", wb.getSheetAt(0).getRow(i + 1).getCell(1).getStringCellValue());
                       ResultSet result = stat.executeQuery(query);
                       while (result.next()) {
                           row.createCell(j + 1).setCellValue(result.getString("id"));
                           row.createCell(j + 2).setCellValue(result.getString("fa"));
                           row.createCell(j + 3).setCellValue(result.getString("im"));
                           row.createCell(j + 4).setCellValue(result.getString("ot"));
                           row.createCell(j + 5).setCellValue(result.getString("npers"));
                           row.createCell(j + 6).setCellValue(sqa.format(sqa1.parse(result.getString("rdat"))));
                           row.createCell(j + 7).setCellValue(result.getString("kat"));
                           pe = 1;
                           f++;
                       }
                       if(pe == 0){
                           int gsp = 0;
                           String query1 = String.format("\n" +
                                   "SELECT t1.id, t1.fa, t1.im, t1.ot, t1.npers, t1.rdat, null as kat, t3.KATLG \n" +
                                   "\tFROM table1 t1\n" +
                                   "left join table3 t3 on t3.id=t1.id\n" +
                                   "where t1.npers='%s' and t1.vibr=1\n" +
                                   "fetch first 1 rows only", wb.getSheetAt(0).getRow(i + 1).getCell(1).getStringCellValue());
                           ResultSet result1 = stat.executeQuery(query1);
                           while (result1.next()) {
                               row.createCell(j + 1).setCellValue(result1.getString("id"));
                               row.createCell(j + 2).setCellValue(result1.getString("fa"));
                               row.createCell(j + 3).setCellValue(result1.getString("im"));
                               row.createCell(j + 4).setCellValue(result1.getString("ot"));
                               row.createCell(j + 5).setCellValue(result1.getString("npers"));
                               row.createCell(j + 6).setCellValue(sqa.format(sqa1.parse(result1.getString("rdat"))));
                               row.createCell(j + 7).setCellValue(result1.getString("katlg"));
                               gsp = 1;
                           }
                           if(gsp==1){
                               f++;
                           }
                       }
                   }
               }

            vyp.close();

            logger.info("запись данных в файл");
            String way = properties.getProperty("way");
            String fileNameForElastic = "itog_"+sqa.format(date)+".xlsx";
            File d = new File(way + fileNameForElastic);
            FileOutputStream file = new FileOutputStream(d);
            xlsx.write(file);
            xlsx.close();

        }
        catch (Exception e) {
            logger.info("connection failed");
            e.printStackTrace();
            logger.error(e.getLocalizedMessage());
        }
        finally{
            if(connection!=null){
                try {
                    connection.close();
                } catch (SQLException e){
                    e.printStackTrace();
                }
            }
        }
    }
}
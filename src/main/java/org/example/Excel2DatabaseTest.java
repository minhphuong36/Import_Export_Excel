package org.example;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Iterator;

public class Excel2DatabaseTest{
    public static void main(String[] args) throws IOException {
        String jdbcURL = "jdbc:mysql://localhost:3306/bbb";
        String username = "root";
        String password = "12345678";

        String excelFilePath = "C:/Users/SMART/Desktop/books.xlsx";

        int batchSize = 20;

        Connection connection = null;

        try {
            long start = System.currentTimeMillis();

            FileInputStream inputStream = new FileInputStream(excelFilePath);

            Workbook workbook = new XSSFWorkbook(inputStream);

            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = firstSheet.iterator();

            connection = DriverManager.getConnection(jdbcURL, username, password);
            connection.setAutoCommit(false);

            String sql = "INSERT INTO book (id, title, price, quantity, money) VALUES (?, ?, ?, ?, ?)";
            PreparedStatement statement = connection.prepareStatement(sql);

            int count = 0;

            rowIterator.next(); // skip the header row

            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell nextCell = cellIterator.next();

                    int columnIndex = nextCell.getColumnIndex();

                    switch (columnIndex) {
                        case 0:
                            int id = (int) nextCell.getNumericCellValue();
                            statement.setInt(1, id);
                            break;
                        case 1:
                            String  title =  nextCell.getStringCellValue();
                            statement.setString(2, title);
                        case 2:
                            double price =  nextCell.getNumericCellValue();
                            statement.setDouble(3, price);
                        case 3:
                            int quantity = (int) nextCell.getNumericCellValue();
                            statement.setInt(4, quantity);
                        case 4:
                            double money =  nextCell.getNumericCellValue();
                            statement.setDouble(5, money);
                    }

                }

                statement.addBatch();

                if (count % batchSize == 0) {
                    statement.executeBatch();
                }

            }

            workbook.close();

            // execute the remaining queries
            statement.executeBatch();

            connection.commit();
            connection.close();

            long end = System.currentTimeMillis();
            System.out.printf("Import done in %d ms\n", (end - start));

        } catch (IOException ex1) {
            System.out.println("Error reading file");
            ex1.printStackTrace();
        } catch (SQLException ex2) {
            System.out.println("Database error");
            ex2.printStackTrace();
        }

    }
}


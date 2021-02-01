package ru.plifis.excelSplit;


import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import java.io.*;
import java.nio.file.Path;
import java.util.LinkedList;
import java.util.List;
import java.util.Scanner;

public class DelimiterExcel {
    /**
     * Делит Excel файл на части по количеству строк
     * @param name имя делимого файла
     * @param rowsOnList количество строк, которое должно быть в выходном файле
     */
        public void delimiter(String name, int rowsOnList) {
            File file = new File(name);
            if (file.exists()) {
                this.readFile(name, rowsOnList);
            } else {
                System.out.println("Файл не найден");
            }
        }

    /**
     * Читает основной (делимый) файл
     * @param name имя основного (делимого) файла
     * @param rowsOnList количество строк, которое должно быть в выходных файлах
     */
        private void readFile(String name, int rowsOnList) {
            try (InputStream in = new FileInputStream(name);
                 HSSFWorkbook bookIn = new HSSFWorkbook(in)) {
                HSSFSheet sheet = bookIn.getSheetAt(0);
                List<HSSFRow> list = new LinkedList<>();
                int numberFile = 0;
                int index = 0;
                int numRows = sheet.getLastRowNum();
                while (index < numRows) {
                    while (list.size() < rowsOnList) {
                        HSSFRow row = sheet.getRow(index);
                        list.add(row);
                        index++;
                    }
                    if (this.writeFile(name, list, numberFile)) {
                        list.clear();
                        numberFile++;
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    /**
     * Записываем строки в новый файл
     * @param name Имя выходного файла (за основу берется имя исходного файла с присоединением порядкового номера)
     * @param list Список строк исхожного файлы, которые записываем в выходной файл
     * @param i порядковый номер файла
     * @return возращаем Истину, если строки записаны в файл
     */
        private boolean writeFile(String name, List<HSSFRow> list, int i) {
            Path newFile = Path.of(i + name);
            try(FileOutputStream out = new FileOutputStream(newFile.toString())) {
             HSSFWorkbook bookOut = new HSSFWorkbook();
            HSSFSheet sheet = bookOut.createSheet();
            for (int indexRow = 0; indexRow < list.size(); indexRow++) {
                HSSFRow oldRow = list.get(indexRow);
                HSSFRow newRow = sheet.createRow(indexRow);
                    try {
                        for (int numberCell = 0; numberCell < 7; numberCell++) {
                            Cell newCell = newRow.createCell(numberCell);
                            Cell oldCell = oldRow.getCell(numberCell);
                                if (oldCell.getCellType() == CellType.STRING) {
                                    String tempStr = oldCell.getStringCellValue();
                                    newCell.setCellValue(tempStr);
                                } else if (oldCell.getCellType() == CellType.NUMERIC) {
                                    double tempDouble = oldCell.getNumericCellValue();
                                    newCell.setCellValue(tempDouble);
                                }
                        }
                    } catch (NullPointerException pointerException) {
                        pointerException.getMessage();
                    }
            }
                bookOut.write(out);
                bookOut.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
            return true;
        }



        public static  void main(String[] args) {
            DelimiterExcel delimiter = new DelimiterExcel();
            //delimiter.delimiter(args[0]);
            Scanner sc = new Scanner(System.in);
            delimiter.delimiter("price.xls", 6000);
        }
    }



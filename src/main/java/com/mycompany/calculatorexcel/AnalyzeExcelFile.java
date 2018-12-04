/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.calculatorexcel;

import static com.mycompany.calculatorexcel.GenerateClassFromString.createBooleanClassOf;
import static com.mycompany.calculatorexcel.GenerateClassFromString.createClassOf;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.xml.transform.stream.StreamSource;
import org.apache.poi.hssf.usermodel.HSSFName;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author Nergal
 */
public class AnalyzeExcelFile {

    private Map<String, Double> table;  //list of numeric values of the cells
    private Map<String, String> tableStr;  //list of String values of the cells
    private Map<String, String> tableNamedCells;
    char[] alphabet = {'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'};

    public String analyze() throws IOException {
        // preparing data for the procedure
        String rezult = "ok ";
        table = new HashMap<String, Double>();
        tableStr = new HashMap<String, String>();
        tableNamedCells = new HashMap<String, String>();

        // use current resource path to get access to the location of calc.xls   
        ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
        StreamSource file = new StreamSource(classLoader.getResourceAsStream("/calc.xls"));
        // формируем из файла экземпляр HSSFWorkbook
        HSSFWorkbook workbook = new HSSFWorkbook(file.getInputStream());
        // выбираем первый лист для обработки 
        // нумерация начинается с 0
        HSSFSheet sheet = workbook.getSheetAt(0);
        // получаем Iterator по всем строкам в листе

        fillTable(sheet); // заполнить таблицу цифровых значений ячеек
        fillTableStr(sheet);
        fillTableNamed(workbook);

        Iterator<Row> rowIterator = sheet.iterator();
        // получаем Iterator по всем ячейкам в строке
        while (rowIterator.hasNext()) {
            Iterator<Cell> cellIterator = rowIterator.next().cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                    if (cell.getCachedFormulaResultType() == Cell.CELL_TYPE_NUMERIC) {
                        System.out.println("NEW Excel formula " + cell.getCellFormula() + " Last evaluated as: " + cell.getNumericCellValue());
                        String f = processStringsOfFormula(processNumbersOfFormula(processNamedCells(cell.getCellFormula())));
                        System.out.println("Formula cells chamged to " + f); // replace cells with known numbers and strings
                        System.out.println("Formula evaluated to " + createClassOf(processFormula(f))); // process formula 
                    }
                    if (cell.getCachedFormulaResultType() == Cell.CELL_TYPE_STRING) {
                        System.out.println("NEW Excel formula " + cell.getCellFormula() + " Last evaluated as: " + cell.getStringCellValue());
                        System.out.println("Formula cells chamged to " + processStringsOfFormula(processNumbersOfFormula(processNamedCells(cell.getCellFormula())))); // replace cells with known numbers and strings
                    }
                }
                String type = converter(cell.toString());
                rezult = rezult + " " + type;
            }
        }
        // данный код позволяет менять excel файл и получать вычисленные значения для ячейки
        // то же самое, что и FormulaEvaluator, делает весь остальной проект
//                    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
//                    // CellValue
//                    HSSFCell cell1 = sheet.getRow(0).getCell(0);
//                    cell1.setCellValue(cell1.getNumericCellValue() +50);
//                    cell.setCellFormula("SUM(A1:A2)");
//                    CellValue cellValue = evaluator.evaluate(cell);
//                    Double value = cellValue.getNumberValue();
//                    System.out.println("evaluator value is "+ value.toString());

        return rezult;
    }

    private String processFormula(String incFormula) {
        System.out.println("Recieved formula for processing " + incFormula);
        String rezult = "rezult= " + incFormula + ";";
        // process expressions in brackets
        Matcher mEXPRSK = Pattern.compile("\\([0-9.\\+\\-\\*\\/]*\\)").matcher(incFormula);
        if (mEXPRSK.find()) {
            rezult = "rezult= (double) " + mEXPRSK.group() + ";";
            System.out.println("EXPRSK-iteration value is " + rezult);
            rezult = processFormula(incFormula.replace(mEXPRSK.group(), createClassOf(rezult).toString()));
        } else {
            // process AND expression
            //Matcher mAND = Pattern.compile("AND\\([^A-Z,]*(,[^A-Z\\),]*)*\\)").matcher(incFormula);
            Matcher mAND = Pattern.compile("AND\\((?!OR|AND|\\()[^,]*(,(?!OR|AND|\\()[^\\\\),]*)*\\)").matcher(incFormula);
            if (mAND.find()) {
                String tempOrFormula = mAND.group().replaceFirst("AND\\(", "");
                tempOrFormula = tempOrFormula.replaceAll(",", " )&( ");
                tempOrFormula = tempOrFormula.replaceAll("=", " == ");
                tempOrFormula = tempOrFormula.replaceAll("<>", " != ");

                rezult = "rezult = (" + tempOrFormula + " ;";
                System.out.println("AND-iteration value is " + rezult);
                rezult = processFormula(incFormula.replace(mAND.group(), String.valueOf(createBooleanClassOf(rezult))));
            } else {
                // process OR expression
                //    Matcher mOR = Pattern.compile("OR\\([^A-Z,]*(,[^A-Z\\),]*)*\\)").matcher(incFormula);
                Matcher mOR = Pattern.compile("OR\\((?!OR|AND|\\()[^,]*(,(?!OR|AND|\\()[^\\\\),]*)*\\)").matcher(incFormula);
                if (mOR.find()) {
                    String tempOrFormula = mOR.group().replaceFirst("OR\\(", "");
                    tempOrFormula = tempOrFormula.replaceAll(",", " )|( ");
                    tempOrFormula = tempOrFormula.replaceAll("=", " == ");
                    tempOrFormula = tempOrFormula.replaceAll("<>", " != ");

                    rezult = "rezult = (" + tempOrFormula + " ;";
                    System.out.println("OR-iteration value is " + rezult);
                    rezult = processFormula(incFormula.replace(mOR.group(), String.valueOf(createBooleanClassOf(rezult))));
                } else {
                    // process if expression
                    //   Pattern pIf = Pattern.compile("IF\\([^\\(\\),]*,[^\\(\\),]*,[^\\(,\\)]*\\)");
                    Matcher mIf = Pattern.compile("IF\\([^\\(\\),]*,[^A-Z,\\(]*,[^A-Z,\\(\\)]*\\)").matcher(incFormula);
                    if (mIf.find()) {
                        String tempFormula = mIf.group();
                        //if (tempFormula.matches("IF\\([^\\=<>A-Z]*=[^\\=,<>]*,[^A-Z,]*,[^A-Z,\\)]*\\)")){
                        if (tempFormula.matches("IF\\([^\\=<>]*=[^\\=,<>]*,[^A-Z,]*,[^A-Z,\\)]*\\)")) {
                            tempFormula = tempFormula.replaceFirst("=", "==");
                        }
                        if (tempFormula.matches("IF\\([^\\=A-Z]*<>[^\\=,]*,[^A-Z,]*,[^A-Z,\\)]*\\)")) {
                            tempFormula = tempFormula.replaceFirst("<>", "!=");
                        }

                        tempFormula = tempFormula.replaceFirst("IF\\(", "");
                        tempFormula = tempFormula.replaceFirst(",", " ){ rezult= (double) ");
                        tempFormula = tempFormula.replaceFirst(",", " ;} else { rezult= (double) ");
                        tempFormula = tempFormula.substring(0, (tempFormula.length() - 1));
                        tempFormula = "if ( " + tempFormula + ";}";
                        System.out.println("IF-iteration value is " + tempFormula);
                        rezult = processFormula(incFormula.replace(mIf.group(), createClassOf(tempFormula).toString()));
                    }
                }
            }
        }

        return rezult;
    }

    // replace named cells with cell addreses
    private String processNamedCells(String cell) {
        String rezult = cell;
        Iterator<Map.Entry<String, String>> it = tableNamedCells.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, String> pair = it.next();
            // search for substrings with cell addreses
            Matcher mCell = Pattern.compile("([^a-zA-Z0-9\"]|^)" + pair.getKey() + "([^a-zA-Z0-9\"]|$)").matcher(rezult);
            while (mCell.find()) {
                String s = mCell.group().replaceFirst(pair.getKey(), pair.getValue());
                rezult = rezult.replace(mCell.group(), s);
            }
        }
        return rezult;
    }

    // replace addreses of numeric cells with numbers
    private String processNumbersOfFormula(String cell) {
        String rezult = cell;
        Iterator<Map.Entry<String, Double>> it = table.entrySet().iterator();

        while (it.hasNext()) {
            Map.Entry<String, Double> pair = it.next();
            // search for substrings with cell addreses
            Matcher mCell = Pattern.compile(pair.getKey() + "\\D").matcher(rezult);
            while (mCell.find()) {
                String s = mCell.group().replaceFirst(pair.getKey(), pair.getValue().toString());
                rezult = rezult.replace(mCell.group(), s);
            }
            rezult = rezult.replaceFirst(pair.getKey() + "$", pair.getValue().toString());
        }
        return rezult;
    }

    // replace addreses of string cells with strings
    private String processStringsOfFormula(String cell) {
        String rezult = cell;
        Iterator<Map.Entry<String, String>> it = tableStr.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, String> pair = it.next();
            Matcher mCell = Pattern.compile(pair.getKey() + "[^0-9]").matcher(rezult);
            while (mCell.find()) {
                String s = mCell.group().substring(0, (mCell.group().length() - 1));
                rezult = rezult.replaceAll(s, "\"" + pair.getValue() + "\"");
            }
        }
        return rezult;
    }

    // get numeric values only and put into table hashMap
    private void fillTable(HSSFSheet sheet) {
        Iterator<Row> rowIterator = sheet.iterator();
        // получаем Iterator по всем ячейкам в строке
        while (rowIterator.hasNext()) {
            Iterator<Cell> cellIterator = rowIterator.next().cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    table.put(String.valueOf(alphabet[cell.getColumnIndex()]) + String.valueOf(cell.getRowIndex() + 1), cell.getNumericCellValue());
                }
            }
        }
    }

    // get numeric values only and put into table hashMap
    private void fillTableStr(HSSFSheet sheet) {
        Iterator<Row> rowIterator = sheet.iterator();
        // получаем Iterator по всем ячейкам в строке
        while (rowIterator.hasNext()) {
            Iterator<Cell> cellIterator = rowIterator.next().cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    tableStr.put(String.valueOf(alphabet[cell.getColumnIndex()]) + String.valueOf(cell.getRowIndex() + 1), cell.getStringCellValue());
                }
            }
        }
    }

    // get named cell values and put into tableNamedCells hashMap
    private void fillTableNamed(HSSFWorkbook wb) {
        for (int i = 0; i < wb.getNumberOfNames(); i++) {
            HSSFName cellName = wb.getNameAt(i);
            System.out.println("cellName is " + cellName.getNameName());
            String form = cellName.getRefersToFormula();
            String addr = form.substring(7, 8) + form.substring(9);
            tableNamedCells.put(cellName.getNameName(), addr);
        }
    }

    private String converter(String cell) {
        String rezult = cell;
        // if there's a number value
        Pattern pNUMB = Pattern.compile("^\\d+(\\.\\d+)?$");
        Matcher mNUMB = pNUMB.matcher(cell);
        if (mNUMB.find()) {
            rezult = cell + "- number";
        }
        // if there's a SUM function
        Pattern pSUM = Pattern.compile("SUM");
        Matcher mSUM = pSUM.matcher(cell);
        if (mSUM.find()) {
            rezult = cell + "- summa";
            System.out.println("summa " + cell);
        }
        return rezult;
    }

}

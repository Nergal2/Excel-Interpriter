/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.calculatorexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.lang.reflect.Method;
import java.net.URI;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.tools.Diagnostic;
import javax.tools.DiagnosticCollector;
import javax.tools.JavaCompiler;
import javax.tools.JavaFileObject;
import javax.tools.SimpleJavaFileObject;
import javax.tools.ToolProvider;
import javax.ws.rs.GET;
import javax.ws.rs.POST;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.QueryParam;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;
import javax.xml.transform.stream.StreamSource;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFName;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author Nergal
 */
@Path("/rest")
public class Restinterface {
    
    private Map<String,Double> table;  //list of numeric values of the cells
    private Map<String,String> tableStr;  //list of String values of the cells
    private Map<String,String> tableNamedCells;
    char[] alphabet  ={ 'A', 'B', 'C', 'D', 'E' , 'F', 'G', 'H' };
    
    @GET
    @Path("/echo")
    public String echo(@QueryParam("q") String original) {
        return original;
    }
    
    @GET
    @Path("/try")
    public String fileTrying() throws IOException {
        // preparing data for the procedure
        String rezult="ok ";
        table = new HashMap<String,Double>();
        tableStr = new HashMap<String,String>();
        tableNamedCells = new HashMap<String,String>();
        
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
                if (cell.getCellType() == Cell.CELL_TYPE_FORMULA){
                    if (cell.getCachedFormulaResultType()==Cell.CELL_TYPE_NUMERIC ){
                        System.out.println("NEW Excel formula " + cell.getCellFormula()+" Last evaluated as: " + cell.getNumericCellValue());
                        String f = processStringsOfFormula(processNumbersOfFormula(processNamedCells(cell.getCellFormula())));
                        System.out.println("Formula cells chamged to " +f); // replace cells with known numbers and strings
                        System.out.println("Formula evaluated to " +createClassOf(processFormula(f))); // process formula 
                    }
                    if (cell.getCachedFormulaResultType()==Cell.CELL_TYPE_STRING ){
                        System.out.println("NEW Excel formula " + cell.getCellFormula()+" Last evaluated as: " + cell.getStringCellValue());
                        System.out.println("Formula cells chamged to " +processStringsOfFormula(processNumbersOfFormula(processNamedCells(cell.getCellFormula())))); // replace cells with known numbers and strings
                    }
               //     System.out.println("NEW Excel formula " + cell.getCellFormula()+" Last evaluated as: " + cell.getNumericCellValue());
              //      System.out.println("Formula cells chamged to " +processStringsOfFormula(processNumbersOfFormula(cell))); // replace cells with known numbers
            //        String form = processNumbersOfFormula(cell);
            //        form = processStringsOfFormula(form);
                    
          //          System.out.println("Formula evaluated to " +createClassOf(processFormula(processNumbersOfFormula(cell)))); // process formula                

//                    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
//                    // CellValue
//                    HSSFCell cell1 = sheet.getRow(0).getCell(0);
//                    cell1.setCellValue(cell1.getNumericCellValue() +50);
//                    cell.setCellFormula("SUM(A1:A2)");
//                    CellValue cellValue = evaluator.evaluate(cell);
//                    Double value = cellValue.getNumberValue();
//                    System.out.println("evaluator value is "+ value.toString());
                }

                
                String type = converter(cell.toString());
                rezult = rezult+" "+type;
        //        if (type.contains("number")) {
        //            table.put(String.valueOf(alplabet[cell.getColumnIndex()])+String.valueOf(cell.getRowIndex()+1), cell.getNumericCellValue());
        //            System.out.println("put " + String.valueOf(alplabet[cell.getColumnIndex()])+String.valueOf(cell.getRowIndex()+1)+ " "+ cell.getNumericCellValue());
        //        }
                
        //        if (type.contains("cell address")) {
        //            System.out.println(cell.toString()+" is "+table.get(cell.toString()));
        //        }
            }
        }
        return rezult;
    }
    
    Double createClassOf(String s){
        //createClass(code);
        JavaCompiler compiler = ToolProvider.getSystemJavaCompiler();
        DiagnosticCollector<JavaFileObject> diagnostics = new DiagnosticCollector<JavaFileObject>();
        StringWriter writer = new StringWriter();
        PrintWriter out = new PrintWriter(writer);
        String code = "public class AnotherClass {";
        //code = code +"  public static void main(String args[]) {";
        code = code +"  public static Double main(String args[]) {";
//        code = code +"    System.out.println(\"This is in another java file\");";
        code = code + "Double rezult ;";
        code = code +  s;
    //    code = code +"    System.out.println(\" value of s "+ s+" is \"+rezult.toString());";
        code = code +"    System.out.println(\" value of s is \"+rezult.toString());";
        code = code +"    return rezult;";
        code = code +"  }";
        code = code +"};";
        out.print(code);
        out.flush();
        out.close();
        JavaFileObject file1 = new JavaSourceFromString("AnotherClass", writer.toString());
        //System.out.println(file1);

        Iterable<? extends JavaFileObject> compilationUnits = Arrays.asList(file1);
        JavaCompiler.CompilationTask task = compiler.getTask(null, null, diagnostics, null, null, compilationUnits);
        boolean success = task.call();
        for (Diagnostic<?> diagnostic : diagnostics.getDiagnostics())
            System.out.println("diagnostic is "+diagnostic);
        System.out.println("Compilation success: " + success);
        
        Double rez = 0.0;
        if (success) {
            try {
                MyClassLoader loader = new MyClassLoader();
                // create class instance
                Class my = loader.getClassFromFile(new File("AnotherClass.class"));
                // use method main
                Method m = my.getMethod("main", new Class[] { String[].class });
                Object o = my.newInstance();
                // call selected method and get the rezult
        //        System.out.println("cacculation rezult is "+m.invoke(o, new Object[] { new String[0] }).toString());
                rez = (Double) m.invoke(o, new Object[] { new String[0] });
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return rez;
    }
    
    private boolean createBooleanClassOf(String s){
        //createClass(code);
        JavaCompiler compiler = ToolProvider.getSystemJavaCompiler();
        DiagnosticCollector<JavaFileObject> diagnostics = new DiagnosticCollector<JavaFileObject>();
        StringWriter writer = new StringWriter();
        PrintWriter out = new PrintWriter(writer);
        String code = "public class AnotherClass {";
        code = code +"  public static boolean main(String args[]) {";
        code = code + "boolean rezult ;";
        code = code +  s;
       // code = code +"    System.out.println(\" value of s "+ s+" is \"+String.valueOf(rezult));";
        code = code +"    System.out.println(\" value of s is \"+String.valueOf(rezult));";
        code = code +"    return rezult;";
        code = code +"  }";
        code = code +"};";
        out.print(code);
        out.flush();
        out.close();
        JavaFileObject file1 = new JavaSourceFromString("AnotherClass", writer.toString());
        
        Iterable<? extends JavaFileObject> compilationUnits = Arrays.asList(file1);
        JavaCompiler.CompilationTask task = compiler.getTask(null, null, diagnostics, null, null, compilationUnits);
        boolean success = task.call();
        for (Diagnostic<?> diagnostic : diagnostics.getDiagnostics())
            System.out.println("diagnostic is "+diagnostic);
        System.out.println("Compilation success: " + success);
        
        boolean rez = false;
        if (success) {
            try {
                MyClassLoader loader = new MyClassLoader();
                // create class instance
                Class my = loader.getClassFromFile(new File("AnotherClass.class"));
                // use method main
                Method m = my.getMethod("main", new Class[] { String[].class });
                Object o = my.newInstance();
                // call selected method and get the rezult
        //        System.out.println("cacculation rezult is "+m.invoke(o, new Object[] { new String[0] }).toString());
                rez = (boolean) m.invoke(o, new Object[] { new String[0] });
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return rez;
    }
    
    static class JavaSourceFromString extends SimpleJavaFileObject {
    final String code;
 
        JavaSourceFromString(String name, String code) {
            super(URI.create("string:///" + name.replace('.', '/') + JavaFileObject.Kind.SOURCE.extension), JavaFileObject.Kind.SOURCE);
            this.code = code;
        }
 
        @Override
        public CharSequence getCharContent(boolean ignoreEncodingErrors) {
            return code;
        }
    }
    
    static class MyClassLoader extends ClassLoader {

        public Class getClassFromFile(File f) {
            byte[] raw = new byte[(int) f.length()];
            System.out.println(f.length());
            InputStream in = null;
            try {
                in = new FileInputStream(f);
                in.read(raw);
            } catch (Exception e) {
                e.printStackTrace();
            }
            try {
                if (in != null)
                    in.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            return defineClass(null, raw, 0, raw.length);
        }
    }
    
    // recursive function for evaluation of formula
    private String processFormula(String incFormula){
        System.out.println("Recieved formula for processing " +incFormula);
        String rezult = "rezult= "+ incFormula + ";";
                // process expressions in brackets
        Matcher mEXPRSK = Pattern.compile("\\([0-9.\\+\\-\\*\\/]*\\)").matcher(incFormula);
        if (mEXPRSK.find()){
            rezult =  "rezult= (double) "+ mEXPRSK.group()+";";
            System.out.println("EXPRSK-iteration value is "+ rezult);
            rezult = processFormula(incFormula.replace(mEXPRSK.group(), createClassOf(rezult).toString()));
        } else {

//        Pattern pEXPR = Pattern.compile("[0-9.]*[\\*\\/][0-9.]*");
//        Matcher mEXPR = pEXPR.matcher(incFormula);
//        if (mEXPR.find()){
//            rezult =  "rezult= "+ mEXPR.group()+";";
//            System.out.println("EXPR-iteration value is "+ rezult);
//            rezult = processFormula(incFormula.replace(mEXPR.group(), createClassOf(rezult).toString()));
//        } else        
//        {
//            Pattern pEXPR2 = Pattern.compile("[0-9.]*[\\+\\-][0-9.]*");
//            Matcher mEXPR2 = pEXPR2.matcher(incFormula);
//            if (mEXPR2.find()){
//                rezult =  "rezult= "+ mEXPR.group()+";";
//                System.out.println("EXPR-iteration value is "+ rezult);
//                rezult = processFormula(incFormula.replace(mEXPR2.group(), createClassOf(rezult).toString()));
//            }
//        }

            // process AND expression
            //Matcher mAND = Pattern.compile("AND\\([^A-Z,]*(,[^A-Z\\),]*)*\\)").matcher(incFormula);
            Matcher mAND = Pattern.compile("AND\\((?!OR|AND|\\()[^,]*(,(?!OR|AND|\\()[^\\\\),]*)*\\)").matcher(incFormula);
            if (mAND.find()){
                String tempOrFormula = mAND.group().replaceFirst("AND\\(", "");
                tempOrFormula = tempOrFormula.replaceAll(",", " )&( ");
                tempOrFormula = tempOrFormula.replaceAll("=", " == ");
                tempOrFormula = tempOrFormula.replaceAll("<>", " != ");

                rezult = "rezult = ("+tempOrFormula+ " ;";
                System.out.println("AND-iteration value is "+ rezult);
                rezult = processFormula(incFormula.replace(mAND.group(), String.valueOf(createBooleanClassOf(rezult))));
            }
            else {

            // process OR expression
        //    Matcher mOR = Pattern.compile("OR\\([^A-Z,]*(,[^A-Z\\),]*)*\\)").matcher(incFormula);
            Matcher mOR = Pattern.compile("OR\\((?!OR|AND|\\()[^,]*(,(?!OR|AND|\\()[^\\\\),]*)*\\)").matcher(incFormula);
            if (mOR.find()){
                String tempOrFormula = mOR.group().replaceFirst("OR\\(", "");
                tempOrFormula = tempOrFormula.replaceAll(",", " )|( ");
                tempOrFormula = tempOrFormula.replaceAll("=", " == ");
                tempOrFormula = tempOrFormula.replaceAll("<>", " != ");

                rezult = "rezult = ("+tempOrFormula+ " ;";
                System.out.println("OR-iteration value is "+ rezult);
                rezult = processFormula(incFormula.replace(mOR.group(), String.valueOf(createBooleanClassOf(rezult))));
            }
            else {

                // process if expression
                //   Pattern pIf = Pattern.compile("IF\\([^\\(\\),]*,[^\\(\\),]*,[^\\(,\\)]*\\)");
                Matcher mIf = Pattern.compile("IF\\([^\\(\\),]*,[^A-Z,\\(]*,[^A-Z,\\(\\)]*\\)").matcher(incFormula);
                if (mIf.find()){
                    String tempFormula = mIf.group();
                    //if (tempFormula.matches("IF\\([^\\=<>A-Z]*=[^\\=,<>]*,[^A-Z,]*,[^A-Z,\\)]*\\)")){
                    if (tempFormula.matches("IF\\([^\\=<>]*=[^\\=,<>]*,[^A-Z,]*,[^A-Z,\\)]*\\)")){
                        tempFormula =  tempFormula.replaceFirst("=", "==");
                    }
                    if (tempFormula.matches("IF\\([^\\=A-Z]*<>[^\\=,]*,[^A-Z,]*,[^A-Z,\\)]*\\)")){
                        tempFormula =  tempFormula.replaceFirst("<>", "!=");
                    }

                    tempFormula =  tempFormula.replaceFirst("IF\\(", "");       
                    tempFormula =  tempFormula.replaceFirst(",", " ){ rezult= (double) ");
                    tempFormula = tempFormula.replaceFirst(",", " ;} else { rezult= (double) ");
                    tempFormula = tempFormula.substring(0, (tempFormula.length()-1));
                    tempFormula = "if ( " + tempFormula+ ";}";
                    System.out.println("IF-iteration value is "+ tempFormula);
                    rezult = processFormula(incFormula.replace(mIf.group(), createClassOf(tempFormula).toString()));
                }
            }
        }}
        return rezult;
    }
    
//    private String processBooleanFormula(String incFormula){
//        System.out.println("Recieved boolean expression for processing " +incFormula);
//        String rezult = "rezult= "+ incFormula + ";";
//      //  boolean rezu =  (1.0>3.0 )|( 5.0>6.0 )|( 8.0 != 8.0 )|( 4.0 == 4.0) ; 
////      Double rr ;
////      if ( 5.0==2 ){ rr=(double) 5.0 ;} else { rr= (double) 1;};
////      rr= 1;
//      
//        return "rezult= "+ createClassOf(rezult).toString()+";";
//        
//    }
    
    
    // replace named cells with cell addreses
    private String processNamedCells(String cell){
        String rezult = cell;
        Iterator<Map.Entry<String, String>> it = tableNamedCells.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, String> pair = it.next();
            // search for substrings with cell addreses
            Matcher mCell = Pattern.compile("([^a-zA-Z0-9\"]|^)"+pair.getKey()+"([^a-zA-Z0-9\"]|$)").matcher(rezult);
            while (mCell.find()){
                String s = mCell.group().replaceFirst(pair.getKey(), pair.getValue());
            //    System.out.println("replace number "+ mCell.group() +" with "+ s);
                rezult = rezult.replace(mCell.group(), s);
            }
            //rezult = rezult.replaceFirst(pair.getKey()+"$", pair.getValue().toString()); 
        }
        return rezult;
    }
    
    
    // replace addreses of numeric cells with numbers
    private String processNumbersOfFormula(String cell){
        //String rezult = cell.getCellFormula();
        String rezult = cell;
        Iterator<Map.Entry<String, Double>> it = table.entrySet().iterator();
        
        while (it.hasNext()) {
            Map.Entry<String, Double> pair = it.next();
            // search for substrings with cell addreses
            Matcher mCell = Pattern.compile(pair.getKey()+"\\D").matcher(rezult);
            while (mCell.find()){
                String s = mCell.group().replaceFirst(pair.getKey(), pair.getValue().toString());
            //    System.out.println("replace number "+ mCell.group() +" with "+ s);
                rezult = rezult.replace(mCell.group(), s);
            }
            rezult = rezult.replaceFirst(pair.getKey()+"$", pair.getValue().toString()); 
        }
        return rezult;
    }
    
     // replace addreses of string cells with strings
    private String processStringsOfFormula(String cell){
        String rezult = cell;
        Iterator<Map.Entry<String, String>> it = tableStr.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, String> pair = it.next();
            Matcher mCell = Pattern.compile(pair.getKey()+"[^0-9]").matcher(rezult);
            while (mCell.find()){
                String s = mCell.group().substring(0, (mCell.group().length()-1));
                rezult= rezult.replaceAll(s, "\""+ pair.getValue()+"\"");
            //    System.out.println("replace string "+ mCell.group() +" with "+ pair.getValue());
            }
            //rezult= rezult.replaceAll(pair.getKey(), "\""+ pair.getValue()+"\"");
        }
        return rezult;
    }
    
    // get numeric values only and put into table hashMap
    private void fillTable (HSSFSheet sheet){
        Iterator<Row> rowIterator = sheet.iterator();
        // получаем Iterator по всем ячейкам в строке
        while (rowIterator.hasNext()) {
            Iterator<Cell> cellIterator = rowIterator.next().cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
                    table.put(String.valueOf(alphabet[cell.getColumnIndex()])+String.valueOf(cell.getRowIndex()+1), cell.getNumericCellValue());
                //    System.out.println("put " + String.valueOf(alphabet[cell.getColumnIndex()])+String.valueOf(cell.getRowIndex()+1)+ " "+ cell.getNumericCellValue());
                }
            }
        }
    }
    
        // get numeric values only and put into table hashMap
    private void fillTableStr (HSSFSheet sheet){
        Iterator<Row> rowIterator = sheet.iterator();
        // получаем Iterator по всем ячейкам в строке
        while (rowIterator.hasNext()) {
            Iterator<Cell> cellIterator = rowIterator.next().cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if(cell.getCellType() == Cell.CELL_TYPE_STRING){
                    tableStr.put(String.valueOf(alphabet[cell.getColumnIndex()])+String.valueOf(cell.getRowIndex()+1), cell.getStringCellValue());
                //    System.out.println("put " + String.valueOf(alphabet[cell.getColumnIndex()])+String.valueOf(cell.getRowIndex()+1)+ " "+ cell.getStringCellValue());
                }
            }
        }
    }
    
     // get named cell values and put into tableNamedCells hashMap
    private void fillTableNamed (HSSFWorkbook wb){
        for (int i=0; i<wb.getNumberOfNames(); i++){
            HSSFName cellName = wb.getNameAt(i);
            System.out.println("cellName is "+cellName.getNameName());
            String form = cellName.getRefersToFormula();
            String addr = form.substring(7, 8) + form.substring(9);
            tableNamedCells.put(cellName.getNameName(), addr);
        }
    }
    
    private String converter (String cell){
        String rezult=cell;
        
        // if there's a cell address
//        Pattern pCELL = Pattern.compile("^[A_Z]\\d$");
//        Matcher mCELL = pCELL.matcher(cell);
//        if (mCELL.find()){
//            rezult =  cell+"- cell address";
//            System.out.println("cell address "+ cell);
//        }        
        
        // if there's a number value
        Pattern pNUMB = Pattern.compile("^\\d+(\\.\\d+)?$");
        Matcher mNUMB = pNUMB.matcher(cell);
        if (mNUMB.find()){
            rezult =  cell+"- number";
        //    System.out.println("number "+ cell);
        }
        
        // if there's a SUM function
        Pattern pSUM = Pattern.compile("SUM");
        Matcher mSUM = pSUM.matcher(cell);
        if (mSUM.find()){
            rezult =  cell+"- summa" ;
            System.out.println("summa "+ cell);
          //  FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
           // CellValue cellValue = evaluator.evaluate(cell);
        }
        
       
        return rezult;
    }
    
    /**
     * set excel file with request
     *
     * @param path - path
     */    
    @POST
    @Path("/file")
    @Produces({MediaType.TEXT_PLAIN})
//    @Consumes(MediaType.APPLICATION_JSON)    
    public Response setFile(String path){
        String result = "ok";
        return Response.ok(result).build();
    }
}

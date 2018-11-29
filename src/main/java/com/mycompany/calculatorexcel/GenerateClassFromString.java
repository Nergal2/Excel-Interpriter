/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.calculatorexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.lang.reflect.Method;
import java.net.URI;
import java.util.Arrays;
import javax.tools.Diagnostic;
import javax.tools.DiagnosticCollector;
import javax.tools.JavaCompiler;
import javax.tools.JavaFileObject;
import javax.tools.SimpleJavaFileObject;
import javax.tools.ToolProvider;

/**
 *
 * @author Nergal
 */
public class GenerateClassFromString {
    public static Double createClassOf(String s){
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
    
    public static boolean createBooleanClassOf(String s){
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
}

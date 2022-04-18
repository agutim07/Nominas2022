/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package excelManager;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.util.Map;
import java.util.Map.Entry;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.ArrayList;

import java.text.DateFormat;  
import java.text.SimpleDateFormat;  
import java.util.Date;  
import java.math.BigInteger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType; 
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;

import org.w3c.dom.Element;
import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

/**
 *
 * @author algut
 */
public class ExcelManager {
    private static MissingCellPolicy xRow;
    
     public static void main(String[] args) throws IOException, TransformerException, ParserConfigurationException {
        //IMPORTAMOS EXCEL CON DATOS
        FileInputStream fis = new FileInputStream(".\\resources\\SistemasInformacionII.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Hoja5");
        int rows = sheet.getLastRowNum();
        
        //CREACION DE DOCUMENTO ERRORES y ERRORESCCC
        DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
        
        Document doc = docBuilder.newDocument();
        Element rootElement = doc.createElement("Trabajadores");
        doc.appendChild(rootElement);
        
        Document docCCC = docBuilder.newDocument();
        Element rootElementCCC = docCCC.createElement("Cuentas");
        docCCC.appendChild(rootElementCCC);
        
        int keySinDNI=0;
        HashMap<String,List<String>> data = new HashMap<String,List<String>>();
        
        //LIST VALUES POSICIONES: 0 NOMBRE, 1 APELLIDO1, 2 APELLIDO2, 3 CIF_EMP
        //4 NOMBRE_EMP, 5 FECHAALTA_EMP, 6 CATEGORIA, 7 PRORRATA, 8 CCC, 9 PAIS CCC
        //10 IBAN, 11 EMAIL
        for(int r=1; r<=rows; r++){
            boolean duplicate=false;
            String falseCCC = "";
            ArrayList<String> listValues = new ArrayList<String>();
            
            if(sheet.getRow(r)!=null && (!checkEmpty(sheet,r,0) || !checkEmpty(sheet,r,1))){
                String key="";
                //COMPROBAMOS SI HAY DNI O ES BLANCO
                if(!checkEmpty(sheet,r,0)){
                    key = sheet.getRow(r).getCell(0).getStringCellValue();
                    String realDNI = calcularLetraDNI(key.substring(0,key.length()-1));
                    if(!key.equals(realDNI)){
                        key = realDNI;
                        sheet.getRow(r).getCell(0).setCellValue(key);
                    }
                    if(checkExists(key,data)) duplicate=true;
                }else{
                    duplicate=true;
                    key="sinDNI_"+keySinDNI;
                    keySinDNI++;
                }
                
                //AÑADIMOS LOS VALORES
                for(int i=1; i<=12; i++){
                    String value;
                    if(!checkEmpty(sheet,r,i)){
                        if(i==6){
                            Date date=sheet.getRow(r).getCell(i).getDateCellValue();
                            DateFormat dateFormat = new SimpleDateFormat("dd-mm-yyyy");  
                            value = dateFormat.format(date);  
                        }else{ 
                            value=sheet.getRow(r).getCell(i).getStringCellValue(); 
                            if(i==9){
                                String realCCC = calcularCCC(value);
                                if(!realCCC.equals(value)){
                                    falseCCC = value;
                                    value = realCCC;
                                    sheet.getRow(r).getCell(i).setCellValue(value);
                                }
                            }
                        }

                        listValues.add(value);
                    }else{
                        String add = "";
                        if(i==12){
                            add = generateEmail(listValues,data);
                            sheet.getRow(r).createCell(12);
                            sheet.getRow(r).getCell(12).setCellValue(add);
                        }
                        if(i==11){
                            add = generarIBAN(listValues.get(8),listValues.get(9));
                            sheet.getRow(r).createCell(11);
                            sheet.getRow(r).getCell(11).setCellValue(add);
                        }
                        listValues.add(add);
                    }
                }
                data.put(key,listValues);
            }
            
            if(duplicate){
                addDuplicate(doc,rootElement,r,listValues);
            }
            if(falseCCC!=""){
                addCCCError(docCCC,rootElementCCC,r,listValues,falseCCC);
            }
            
        }
        
        //for(Map.Entry entry:data.entrySet()){
        //    System.out.println(entry.getKey());
        //}
      
        createErroresFile(doc,"dni");
        createErroresFile(docCCC,"ccc");
        FileOutputStream out = new FileOutputStream(new File(".\\resources\\SistemasInformacionII_Actualizado.xlsx"));
        workbook.write(out);
        out.close();
      
        
    }
     
    public static boolean checkExists(String key, HashMap<String,List<String>> data){
        key = key.substring(0,key.length()-1);
        
        for(Map.Entry entry:data.entrySet()){
            String key2 = entry.getKey().toString().substring(0,key.length());
            if(key.equals(key2)) return true;
        }
        
        return false;
    }
    
    public static String generateEmail(List<String> listValues, HashMap<String,List<String>> data){
        String email="";
        if(listValues.get(2)!="")email+=listValues.get(2).charAt(0);
        email+=listValues.get(1).charAt(0);
        email+=listValues.get(0).charAt(0);
        String parte2 = "@"+listValues.get(4)+".es";
        int ap = checkEmailExists(email+parte2,data);
        String num="";
        if(ap<10){
            num = 0 +  String.valueOf(ap);
        }else{
            num = String.valueOf(ap);
        }
        
        return email+num+parte2;
    }
    
    public static int checkEmailExists(String email, HashMap<String,List<String>> data){
        int apariciones=0;
        
        for(Map.Entry entry:data.entrySet()){
            String key = entry.getKey().toString();
            String email2 = data.get(key).get(11);
            String email3="";
            for (int i=0; i<email2.length(); i++) { 
                char c = email2.charAt(i);
                boolean añadir = true;
                if((i+2)<email2.length()){
                    if(email2.charAt(i+1)=='@')añadir=false;
                    if(email2.charAt(i+2)=='@')añadir=false;
                }
                if(añadir) email3+=c;
            }
            if(email.equals(email3)) apariciones++;
        }
        
        return apariciones;
    }
    
    
    
    public static void addDuplicate(Document doc, Element rootElement, int r, ArrayList<String> listValues){
        Element elemento1 = doc.createElement("Trabajador");
        Attr attr = doc.createAttribute("id");
        attr.setValue(String.valueOf(r+1));
        elemento1.setAttributeNode(attr);
        Element elementoNom = doc.createElement("Nombre");
        elementoNom.setTextContent(listValues.get(0));
        elemento1.appendChild(elementoNom);
        Element elementoAp1 = doc.createElement("PrimerApellido");
        elementoAp1.setTextContent(listValues.get(1));
        elemento1.appendChild(elementoAp1);
        Element elementoAp2 = doc.createElement("SegundoApellido");
        elementoAp2.setTextContent(listValues.get(2));
        elemento1.appendChild(elementoAp2);
        Element elementoEmp = doc.createElement("Empresa");
        elementoEmp.setTextContent(listValues.get(4));
        elemento1.appendChild(elementoEmp);
        Element elementoCat = doc.createElement("Categoria");
        elementoCat.setTextContent(listValues.get(6));
        elemento1.appendChild(elementoCat);
        rootElement.appendChild(elemento1);
    }
    
    public static void addCCCError(Document doc, Element rootElement, int r, ArrayList<String> listValues, String falseCCC){
        Element elemento1 = doc.createElement("Cuenta");
        Attr attr = doc.createAttribute("id");
        attr.setValue(String.valueOf(r+1));
        elemento1.setAttributeNode(attr);
        Element elementoNom = doc.createElement("Nombre");
        elementoNom.setTextContent(listValues.get(0));
        elemento1.appendChild(elementoNom);
        Element elementoAp = doc.createElement("Apellidos");
        elementoAp.setTextContent(listValues.get(1)+" "+listValues.get(2));
        elemento1.appendChild(elementoAp);
        Element elementoEmp = doc.createElement("Empresa");
        elementoEmp.setTextContent(listValues.get(4));
        elemento1.appendChild(elementoEmp);
        Element elementoCcc = doc.createElement("CCCErroneo");
        elementoCcc.setTextContent(falseCCC);
        elemento1.appendChild(elementoCcc);
        Element elementoIBAN = doc.createElement("IBANCorrecto");
        elementoIBAN.setTextContent(listValues.get(10));
        elemento1.appendChild(elementoIBAN);
        rootElement.appendChild(elemento1);
    }
    
    public static String calcularCCC(String originalNum){
        if(originalNum.length()!=20){
            return "";
        }
        
        String e = originalNum.substring(0,4);
        String o = originalNum.substring(4,8);
        String n = originalNum.substring(10);
        
        String controld1 = "00"+e+o;
        String controld2 = n;
        int sumd1=0, sumd2=0;
        
        for(int i=0; i<10; i++){
            int factor = (int) Math.pow(2, i) % 11;
            sumd1+=Character.getNumericValue(controld1.charAt(i)) * factor;  
            sumd2+=Character.getNumericValue(controld2.charAt(i)) * factor;
        }
        
        int d1 = 11 - (sumd1%11);
        int d2 = 11 - (sumd2%11);
        if(d1==11) d1=0;
        if(d1==10) d1=1;
        if(d2==11) d2=0;
        if(d2==10) d2=1;
        
        String d = String.valueOf(d1) + String.valueOf(d2);
        
        return e+o+d+n;
    }
    
    public static String generarIBAN(String ccc, String pais){
        if(ccc.length()!=20 || pais.length()!=2){
            return "";
        }
        
        char[] paisChar = {pais.charAt(0), pais.charAt(1)};
        int[] paisDigits = new int[2];
        
        for(int i=0; i<2; i++){
            switch(paisChar[i]){
                case 'A': paisDigits[i]=10; break;
                case 'B': paisDigits[i]=11; break;
                case 'C': paisDigits[i]=12; break;
                case 'D': paisDigits[i]=13; break;
                case 'E': paisDigits[i]=14; break;
                case 'F': paisDigits[i]=15; break;
                case 'G': paisDigits[i]=16; break;
                case 'H': paisDigits[i]=17; break;
                case 'I': paisDigits[i]=18; break;
                case 'J': paisDigits[i]=19; break;
                case 'K': paisDigits[i]=20; break;
                case 'L': paisDigits[i]=21; break;
                case 'M': paisDigits[i]=22; break;
                case 'N': paisDigits[i]=23; break;
                case 'O': paisDigits[i]=24; break;
                case 'P': paisDigits[i]=25; break;
                case 'Q': paisDigits[i]=26; break;
                case 'R': paisDigits[i]=27; break;
                case 'S': paisDigits[i]=28; break;
                case 'T': paisDigits[i]=29; break;
                case 'U': paisDigits[i]=30; break;
                case 'V': paisDigits[i]=31; break;
                case 'W': paisDigits[i]=32; break;
                case 'X': paisDigits[i]=33; break;
                case 'Y': paisDigits[i]=34; break;
                case 'Z': paisDigits[i]=35; break;
            }
        }
        
        String ibanCal = ccc+String.valueOf(paisDigits[0])+String.valueOf(paisDigits[1])+"00";
        BigInteger ibanNum =  new BigInteger(ibanCal);
        int ibanNumRest = (ibanNum.remainder(BigInteger.valueOf(97))).intValue();
        String dato = String.valueOf(98 - ibanNumRest);
        if(dato.length()==1) dato = "0"+dato;
        
        return pais+dato+ccc;
    }
    
    
    public static String calcularLetraDNI(String originalNum){
        String num = originalNum;
        if(num.substring(0,1).equals("X")) num = "0"+num.substring(1);
        if(num.substring(0,1).equals("Y")) num = "1"+num.substring(1);
        if(num.substring(0,1).equals("Z")) num = "2"+num.substring(1);
        
        int dni = Integer.valueOf(num);
        int resto = dni%23;
        String letra = "";
        switch(resto){
            case 0: letra="T"; break;
            case 1: letra="R"; break;
            case 2: letra="W"; break;
            case 3: letra="A"; break;
            case 4: letra="G"; break;
            case 5: letra="M"; break;
            case 6: letra="Y"; break;
            case 7: letra="F"; break;
            case 8: letra="P"; break;
            case 9: letra="D"; break;
            case 10: letra="X"; break;
            case 11: letra="B"; break;
            case 12: letra="N"; break;
            case 13: letra="J"; break;
            case 14: letra="Z"; break;
            case 15: letra="S"; break;
            case 16: letra="Q"; break;
            case 17: letra="V"; break;
            case 18: letra="H"; break; 
            case 19: letra="L"; break;
            case 20: letra="C"; break;
            case 21: letra="K"; break; 
            case 22: letra="E"; break;
        }
        
        return originalNum+letra;
    }
     
    public static boolean checkEmpty(XSSFSheet sheet, int r, int i){
        Cell cell = sheet.getRow(r).getCell(i, xRow.RETURN_BLANK_AS_NULL);
        if(cell == null || cell.getCellType() == CellType.BLANK){
            return true;
        }
        
        return false;
    }
    
    public static void createErroresFile(Document doc, String type) throws TransformerException, ParserConfigurationException{
      String filepath="";
      if(type=="dni") filepath = ".\\resources\\Errores.xml";
      if(type=="ccc") filepath = ".\\resources\\ErroresCCC.xml";
      TransformerFactory transformerFactory = TransformerFactory.newInstance();
      Transformer transformer = transformerFactory.newTransformer();
      transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
      transformer.setOutputProperty(OutputKeys.INDENT, "yes");
      transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", String.valueOf(2));
      DOMSource source = new DOMSource(doc);
      StreamResult result = new StreamResult(new File(filepath));
      transformer.transform(source, result);
    }
}

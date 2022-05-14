/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package excelManager;

import map.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.File;
import java.text.ParseException;
import java.util.*;
import java.io.IOException;

import java.text.SimpleDateFormat;
import java.math.BigInteger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
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
    public static List<Trabajadorbbdd> data;
    public static List<Empresas> dataEmp;
    public static List<Categorias> dataCat;
    public static List<Nomina> dataNom;
    
     public static void main(String[] args) throws IOException, TransformerException, ParserConfigurationException, ParseException {
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
        data = new ArrayList<>();
        dataEmp = new ArrayList<>();
        dataCat = new ArrayList<>();
        dataNom = new ArrayList<>();

        //EN ESTAS DOS LISTA, LA POS X COINCIDE CON EL DATO DEL TRABAJOR X EN EL ARRAY 'data'
        ArrayList<String> listaProrrateos = new ArrayList<>();
        ArrayList<Integer> listaPosExcel = new ArrayList<>();

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
                    if(checkDNIExists(key)) duplicate=true;
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
                            SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
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
                            add = generateEmail(listValues);
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

                //SI EL TRABAJADOR NO EXISTE LO AÑADIMOS A LA BASE DE DATOS (SI TIENE EL DNI EN BLANCO NO SE INSERTA)
                if(!checkTrabajadorExists(key,listValues.get(0),listValues.get(5))) {
                    Trabajadorbbdd nuevo = getTrabajador(key,listValues);
                    listaProrrateos.add(listValues.get(7));
                    listaPosExcel.add(r+1);
                    data.add(nuevo);
                }
            }

            if(duplicate){
                addDuplicate(doc,rootElement,r,listValues);
            }
            if(falseCCC!=""){
                addCCCError(docCCC,rootElementCCC,r,listValues,falseCCC);
            }

        }

        //GENERADOR DE NOMINAS
        NominasGenerator.generarNominas(data,dataCat,dataNom,listaProrrateos,workbook);
        NominasPDFGenerator.generarPDFs(dataNom);

        //ALMACENAR/ACTUALIZAR LOS DATOS EN LA BD


        //CREACION DE DOCUMENTOS Y FIN
        //Document docNom = getNominasDoc(listaPosExcel,docBuilder);
        createXMLFile(doc,"dni");
        createXMLFile(docCCC,"ccc");
        //createXMLFile(docNom,"nom");
        FileOutputStream out = new FileOutputStream(new File(".\\resources\\SistemasInformacionII_Actualizado.xlsx"));
        workbook.write(out);
        out.close();
      
        
    }

    public static Trabajadorbbdd getTrabajador(String key, ArrayList<String> listValues) throws ParseException {
        //LIST VALUES POSICIONES: 0 NOMBRE, 1 APELLIDO1, 2 APELLIDO2, 3 CIF_EMP
        //4 NOMBRE_EMP, 5 FECHAALTA_EMP, 6 CATEGORIA, 7 PRORRATA, 8 CCC, 9 PAIS CCC
        //10 IBAN, 11 EMAIL
        Trabajadorbbdd t = new Trabajadorbbdd();
        t.setNifnie(key);
        t.setNombre(listValues.get(0));
        t.setApellido1(listValues.get(1));
        t.setApellido2(listValues.get(2));
        t.setFechaAlta(new SimpleDateFormat("dd/MM/yyyy").parse(listValues.get(5)));
        t.setCodigoCuenta(listValues.get(8));
        t.setIban(listValues.get(10));
        t.setEmail(listValues.get(11));

        int posEmpresa = checkEmpresa(listValues.get(3),listValues.get(4));
        t.setEmpresas(dataEmp.get(posEmpresa));

        int posCategoria = checkCategoria(listValues.get(6));
        t.setCategorias(dataCat.get(posCategoria));

        Set empresasCollecion = dataEmp.get(posEmpresa).getTrabajadorbbdds();
        empresasCollecion.add(t);
        dataEmp.get(posEmpresa).setTrabajadorbbdds(empresasCollecion);

        Set categoriaCollecion = dataCat.get(posCategoria).getTrabajadorbbdds();
        categoriaCollecion.add(t);
        dataCat.get(posCategoria).setTrabajadorbbdds(categoriaCollecion);

        return t;
    }

    public static int checkCategoria(String cat){
        //VEMOS SI EXISTE LA CATEGORIA, SINO EXISTE LA CREAMOS
        int i;
        for(i=0; i<dataCat.size(); i++){
            if(dataCat.get(i).getNombreCategoria().equals(cat)){
                return i;
            }
        }

        Categorias nueva = new Categorias();
        nueva.setNombreCategoria(cat);
        dataCat.add(nueva);
        return i;
    }

    public static int checkEmpresa(String cif, String nombre){
        //VEMOS SI EXISTE LA EMPRESA, SINO EXISTE LA CREAMOS
        int i;
        for(i=0; i<dataEmp.size(); i++){
            if(dataEmp.get(i).getCif().equals(cif)){
                return i;
            }
        }

        Empresas nueva = new Empresas();
        nueva.setCif(cif);
        nueva.setNombre(nombre);
        dataEmp.add(nueva);
        return i;
    }



    public static boolean checkTrabajadorExists(String dni, String nombre, String date) throws ParseException {
        //SI EL DNI ESTÁ EN BLANCO NO SE INSERTA
        if(dni.substring(0,6).equals("sinDNI")) return true;

        Date fechaAlta = new SimpleDateFormat("dd/MM/yyyy").parse(date);
        //UN TRABAJADOR EXISTE SI YA EXISTE ALGUIEN CON SU NOMBRE, DNI Y FECHA DE ALTA
        for(int i=0; i<data.size(); i++){
            Trabajadorbbdd iter = data.get(i);
            if(iter.getNifnie().equals(dni) && iter.getNombre().equals(nombre) && iter.getFechaAlta().equals(fechaAlta)){
                return true;
            }
        }

        return false;
    }

    public static boolean checkDNIExists(String key){
        key = key.substring(0,key.length()-1);
        
        for(int i=0; i<data.size(); i++){
            String key2 = data.get(i).getNifnie().substring(0,key.length());
            if(key.equals(key2)) return true;
        }
        
        return false;
    }

    public static int getExcelRow(Trabajadorbbdd trab, ArrayList<Integer> rows){
        //PARA ENCONTRAR AL TRABAJADOR BUSCAMOS UNO CON SU NOMBRE, DNI Y FECHA DE ALTA
        for(int i=0; i<data.size(); i++){
            Trabajadorbbdd iter = data.get(i);
            if(iter.getNifnie().equals(trab.getNifnie()) && iter.getNombre().equals(trab.getNombre()) && iter.getFechaAlta().equals(trab.getFechaAlta())){
                return rows.get(i);
            }
        }

        return -1;
    }
    
    public static String generateEmail(List<String> listValues){
        String email="";
        if(listValues.get(2)!="")email+=listValues.get(2).charAt(0);
        email+=listValues.get(1).charAt(0);
        email+=listValues.get(0).charAt(0);
        String parte2 = "@"+listValues.get(4)+".es";
        int ap = checkEmailExists(email+parte2);
        String num="";
        if(ap<10){
            num = 0 +  String.valueOf(ap);
        }else{
            num = String.valueOf(ap);
        }
        
        return email+num+parte2;
    }
    
    public static int checkEmailExists(String email){
        int apariciones=0;
        
        for(int j=0; j<data.size(); j++){
            String email2 = data.get(j).getEmail();
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

    public static Document getNominasDoc(ArrayList<Integer> rows, DocumentBuilder docBuilder){
        Document doc = docBuilder.newDocument();
        Element rootElement = doc.createElement("Nominas");
        Attr attr1 = doc.createAttribute("fechaNomina");
        attr1.setValue(dataNom.get(0).getMes()+"/"+dataNom.get(0).getAnio());
        rootElement.setAttributeNode(attr1);
        doc.appendChild(rootElement);

        for(int i=0; i<dataNom.size(); i++){
            Nomina n = dataNom.get(i);

            Element elemento1 = doc.createElement("Nomina");
            Attr attr = doc.createAttribute("idNomina");
            attr.setValue(String.valueOf(n.getIdNomina()));
            elemento1.setAttributeNode(attr);

            Element elementoExtra = doc.createElement("Extra");
            String extra = "N";
            if(n.getBaseEmpresario()==0) extra = "S";
            elementoExtra.setTextContent(extra);
            elemento1.appendChild(elementoExtra);

            Element elementoExcel = doc.createElement("idFilaExcel");
            elementoExcel.setTextContent(String.valueOf(getExcelRow(n.getTrabajadorbbdd(),rows)));
            elemento1.appendChild(elementoExcel);

            Element elementoNom = doc.createElement("Nombre");
            elementoNom.setTextContent(n.getTrabajadorbbdd().getNombre());
            elemento1.appendChild(elementoNom);

            Element elementoNIF = doc.createElement("NIF");
            elementoNIF.setTextContent(n.getTrabajadorbbdd().getNifnie());
            elemento1.appendChild(elementoNIF);

            Element elementoIBAN = doc.createElement("IBAN");
            elementoIBAN.setTextContent(n.getTrabajadorbbdd().getIban());
            elemento1.appendChild(elementoIBAN);

            Element elementoCat = doc.createElement("Categoria");
            elementoCat.setTextContent(n.getTrabajadorbbdd().getCategorias().getNombreCategoria());
            elemento1.appendChild(elementoCat);

            Element elementoBA = doc.createElement("BrutoAnual");
            elementoBA.setTextContent(String.valueOf(n.getBrutoAnual()));
            elemento1.appendChild(elementoBA);

            Element elementoIRPF = doc.createElement("ImporteIrpf");
            elementoIRPF.setTextContent(String.valueOf(n.getImporteIrpf()));
            elemento1.appendChild(elementoIRPF);

            Element elementoBEmp = doc.createElement("BaseEmpresario");
            elementoBEmp.setTextContent(String.valueOf(n.getBaseEmpresario()));
            elemento1.appendChild(elementoBEmp);

            Element elementoBN = doc.createElement("BrutoNomina");
            elementoBN.setTextContent(String.valueOf(n.getBrutoNomina()));
            elemento1.appendChild(elementoBN);

            Element elementoLN = doc.createElement("LiquidoNomina");
            elementoLN.setTextContent(String.valueOf(n.getLiquidoNomina()));
            elemento1.appendChild(elementoLN);

            Element elementoCTEmp = doc.createElement("CosteTotalEmpresario");
            elementoCTEmp.setTextContent(String.valueOf(n.getCosteTotalEmpresario()));
            elemento1.appendChild(elementoCTEmp);

            rootElement.appendChild(elemento1);
        }

        return doc;
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
    
    public static void createXMLFile(Document doc, String type) throws TransformerException, ParserConfigurationException{
      String filepath="";
      if(type=="dni") filepath = ".\\resources\\Errores.xml";
      if(type=="ccc") filepath = ".\\resources\\ErroresCCC.xml";
      if(type=="nom") filepath = ".\\resources\\Nominas.xml";
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

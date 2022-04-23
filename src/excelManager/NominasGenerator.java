package excelManager;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.ParseException;
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
import java.util.concurrent.TimeUnit;

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

public class NominasGenerator {
    public static void generarNominas(HashMap<String,List<String>> data, XSSFWorkbook workbook) throws ParseException {
        //CONSEGUIMOS LOS DATOS DE LAS HOJAS 1-4
        ArrayList<Categoria> listaCategorias = getHoja3(workbook);
        ArrayList<BrutoRetenciones> listaRetenciones = getHoja4(workbook);

        ArrayList<Double> listaCuotas = getHoja1(workbook);
            //listaCuotas: 0 seguridad social Trab, 1 desempleo Trab, 2 formacion Trab
            //3 accidentes Emp, 4 contingencias Emp, 5 fogasa Emp, 6 desempleo Emp, 7 formacion Emp

        ArrayList<Integer> listaImporteTrienios = getHoja2(workbook);
            //valor en pos x = importe con x nº de trienios

        //INTRODUCIMOS POR CONSOLA LA FECHA A CALCULAR LAS NOMINAS Y RECORREMOS TODOS LOS TRABAJADORES
        String mainDateString = "01/"+"06/2014";
        Date mainDate = new SimpleDateFormat("dd/MM/yyyy").parse(mainDateString);

        for(Map.Entry entry:data.entrySet()){
            String key = entry.getKey().toString();

            //CHECKEAMOS LA FECHA DE ALTA EN LA EMPRESA CON LA INTRODUCIDA
            String dateAltaString = data.get(key).get(5);
            System.out.print(key + " " + dateAltaString + " ");
            Date dateAlta = new SimpleDateFormat("dd/MM/yyyy").parse(dateAltaString);

            TimeUnit time = TimeUnit.DAYS;
            long daysDiff = time.convert(mainDate.getTime() - dateAlta.getTime(), TimeUnit.MILLISECONDS);

            //SI DAYSDIFF<31 DÍAS NO SE CALCULA LA NOMINA
            if(daysDiff>31){
                int trienio = (Integer.parseInt(mainDateString.substring(6))-Integer.parseInt(dateAltaString.substring(6)))/3;
                System.out.print(trienio + " ");

                String pro = data.get(key).get(7); boolean prorrateo=false;
                if(pro.equals("SI")) prorrateo=true;
                System.out.print(prorrateo + " ");

                String categoria = data.get(key).get(6);
                int[] salariocomp = getSalarioyComplementos(categoria,listaCategorias);
                int salarioBase = salariocomp[0]; int complementos = salariocomp[1];
                Double salarioBaseMes = salarioBase/14.0; Double complementosMes = complementos/14.0;
                int antiguedadMes = listaImporteTrienios.get(trienio); int antiguedadAnual = antiguedadMes*14;

                Double brutoAnual = Double.valueOf(salarioBase+complementos+antiguedadAnual);

                //CASO: NO TRABAJA EL ANO ENTERO
                if(daysDiff<365){
                    Double[] extras = getExtrasRecienContratado(dateAltaString, mainDateString);
                    int nominasCompletas = extras[0].intValue();
                    Double porcentajeExtraDic = extras[1]; Double porcentajeExtraJun = extras[2];

                    brutoAnual = (salarioBaseMes + complementosMes) * nominasCompletas;

                    if(!prorrateo){
                        brutoAnual += (salarioBaseMes + complementosMes) * (porcentajeExtraDic + porcentajeExtraJun);
                    }
                }else{
                    //CASO: PRORRATEO Y CAMBIO DE TRIENIO AL SIGUIENTE AÑO
                    int nextTrienio = ((Integer.parseInt(mainDateString.substring(6))+1)-Integer.parseInt(dateAltaString.substring(6)))/3;
                    if(prorrateo && nextTrienio!=trienio){
                        brutoAnual += (double) listaImporteTrienios.get(nextTrienio)/6 - (double) antiguedadMes/6;
                    }
                }
                System.out.print(prec(brutoAnual)+" ");
                //FIN CALCULO DE BRUTO ANUAL
                Double irpf = getIRPF(listaRetenciones, brutoAnual);
                Double prorrateoExtra = salarioBaseMes/6 + complementosMes/6 + (double) antiguedadMes/6;
                if(prorrateo) System.out.print("-- "+prec(prorrateoExtra));






            }
            System.out.println();
        }
    }

    public static int[] getSalarioyComplementos(String cat, ArrayList<Categoria> list){
        for(int i=0; i< list.size(); i++){
            if(cat.equals(list.get(i).getCategoria())){
                return new int[]{list.get(i).getSalarioBase(),list.get(i).getComplementos()};
            }
        }

        return new int[]{0,0};
    }

    public static Double[] getExtrasRecienContratado(String dateAlta, String mainDate){
        int mesAlta = Integer.valueOf(dateAlta.substring(3,5));
        int añoAlta = Integer.parseInt(mainDate.substring(6));
        int añoActual = Integer.parseInt(dateAlta.substring(6));

        int nominasCompletas=12;
        int extraDic=6;
        int extraJun=6;

        if(añoActual==añoAlta){
            nominasCompletas = 12-mesAlta + 1;
            if(mesAlta>6){
                extraDic = 12-mesAlta;
            }
            extraJun=0;
            if(mesAlta<6){
                extraJun = 12-mesAlta-6;
            }
        }

        Double porcentajeExtraDiciembre = extraDic/6.0;
        Double porcentajeExtraJunio = extraJun/6.0;

        return new Double[]{Double.valueOf(nominasCompletas),porcentajeExtraDiciembre,porcentajeExtraJunio};
    }

    public static Double getIRPF(ArrayList<BrutoRetenciones> lista, Double bruto){
        for(int i=0; i<(lista.size()-1); i++){
            if(bruto>=lista.get(i).getBrutoAnual() && bruto<lista.get(i+1).getBrutoAnual()){
                return lista.get(i).getRetencion();
            }
        }

        return 0.0;
    }

    public static ArrayList<Categoria> getHoja3(XSSFWorkbook workbook){
        //INTRODUCIMOS HOJA 3: CATEGORIA, COMPLEMENTO Y SALARIO BASE
        XSSFSheet sheet = workbook.getSheet("Hoja3");
        int rows = sheet.getLastRowNum();
        ArrayList<Categoria> lista = new ArrayList<>();

        for(int r=1; r<=rows; r++) {
            String cat = sheet.getRow(r).getCell(0).getStringCellValue();
            int comp = (int) sheet.getRow(r).getCell(1).getNumericCellValue();
            int sal = (int) sheet.getRow(r).getCell(2).getNumericCellValue();
            lista.add(new Categoria(cat,sal,comp));
        }

        return lista;
    }

    public static ArrayList<BrutoRetenciones> getHoja4(XSSFWorkbook workbook){
        //INTRODUCIMOS HOJA 4: BRUTO ANUAL Y RETENCIONES
        XSSFSheet sheet = workbook.getSheet("Hoja4");
        int rows = sheet.getLastRowNum();
        ArrayList<BrutoRetenciones> lista = new ArrayList<>();

        for(int r=1; r<=rows; r++) {
            int bruto = (int) sheet.getRow(r).getCell(0).getNumericCellValue();
            Double retencion = sheet.getRow(r).getCell(1).getNumericCellValue();
            lista.add(new BrutoRetenciones(bruto,retencion));
        }

        return lista;
    }

    public static ArrayList<Double> getHoja1(XSSFWorkbook workbook){
        //INTRODUCIMOS HOJA 1
        XSSFSheet sheet = workbook.getSheet("Hoja1");
        int rows = sheet.getLastRowNum();
        ArrayList<Double> lista = new ArrayList<>();

        for(int r=0; r<=rows; r++) {
            double cuota =  sheet.getRow(r).getCell(1).getNumericCellValue();
            lista.add(cuota);
        }

        return lista;
    }

    public static ArrayList<Integer> getHoja2(XSSFWorkbook workbook){
        //INTRODUCIMOS HOJA 2
        XSSFSheet sheet = workbook.getSheet("Hoja2");
        int rows = sheet.getLastRowNum();
        ArrayList<Integer> lista = new ArrayList<>();
        lista.add(0); //LA POSICION 0, =0 TRIENIOS TIENE UN IMPORTE BRUTO DE 0

        for(int r=1; r<=rows; r++) {
            int importe =  (int) sheet.getRow(r).getCell(1).getNumericCellValue();
            lista.add(importe);
        }

        return lista;
    }

    public static class Categoria{
        private String cat;
        private int salarioBase;
        private int complementos;

        Categoria(String c, int s, int comp){
            this.cat = c;
            this.salarioBase = s;
            this.complementos = comp;
        }

        public String getCategoria(){ return this.cat; }
        public int getSalarioBase(){ return this.salarioBase; }
        public int getComplementos(){ return this.complementos; }
    }

    public static class BrutoRetenciones{
        private int bruto;
        private Double retencion;

        BrutoRetenciones(int b, Double r){
            this.bruto = b;
            this.retencion = r;
        }

        public int getBrutoAnual(){ return this.bruto; }
        public Double getRetencion(){ return this.retencion; }
    }

    public static Double prec(Double x){
        Double newDouble = BigDecimal.valueOf(x)
                .setScale(2, RoundingMode.HALF_UP)
                .doubleValue();

        return newDouble;
    }
}

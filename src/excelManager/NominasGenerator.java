package excelManager;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.net.MalformedURLException;
import java.text.ParseException;
import java.util.*;

import java.text.SimpleDateFormat;
import java.util.concurrent.TimeUnit;

import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.properties.TextAlignment;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;

import org.w3c.dom.Element;
import org.w3c.dom.Attr;
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
    public static void generarNominas(HashMap<String,List<String>> data, XSSFWorkbook workbook) throws ParseException, IOException {
        //CONSEGUIMOS LOS DATOS DE LAS HOJAS 1-4
        ArrayList<Categoria> listaCategorias = getHoja3(workbook);
        ArrayList<BrutoRetenciones> listaRetenciones = getHoja4(workbook);

        ArrayList<Double> listaCuotas = getHoja1(workbook);
            //listaCuotas: 0 seguridad social Trab, 1 desempleo Trab, 2 formacion Trab
            //3 accidentes Emp, 4 seguridad social Emp, 5 fogasa Emp, 6 desempleo Emp, 7 formacion Emp

        ArrayList<Integer> listaImporteTrienios = getHoja2(workbook);
            //valor en pos x = importe con x nº de trienios

        //INTRODUCIMOS POR CONSOLA LA FECHA A CALCULAR LAS NOMINAS Y RECORREMOS TODOS LOS TRABAJADORES
        Scanner lectura = new Scanner(System.in);
        System.out.println("Introduce fecha en la cual generar las nóminas (mm/aaaa): ");
        String dateEntrada = lectura.next();

        String mainDateString = "01/"+dateEntrada;
        Date mainDate = new SimpleDateFormat("dd/MM/yyyy").parse(mainDateString);

        int prueba=1;

        for(Map.Entry entry:data.entrySet()){
            if(prueba==1){
                PDFgenerator(entry.getKey().toString(),data);
            }
            prueba++;

            String key = entry.getKey().toString();

            //CHECKEAMOS LA FECHA DE ALTA EN LA EMPRESA CON LA INTRODUCIDA
            String dateAltaString = data.get(key).get(5);
            Date dateAlta = new SimpleDateFormat("dd/MM/yyyy").parse(dateAltaString);

            TimeUnit time = TimeUnit.DAYS;
            long daysDiff = time.convert(mainDate.getTime() - dateAlta.getTime(), TimeUnit.MILLISECONDS);

            //SI DAYSDIFF<28 DÍAS NO SE CALCULA LA NOMINA
            if(daysDiff>=28){
                int añoCalculo = Integer.parseInt(mainDateString.substring(6));
                int trienio = (añoCalculo-Integer.parseInt(dateAltaString.substring(6)))/3;
                int mesCalculo = Integer.valueOf(mainDateString.substring(3,5));

                String pro = data.get(key).get(7); boolean prorrateo=false;
                if(pro.equals("SI")) prorrateo=true;

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
                //FIN CALCULO DE BRUTO ANUAL

                Double irpfRetencion = getIRPF(listaRetenciones, brutoAnual);
                Double prorrateoExtra = salarioBaseMes/6 + complementosMes/6 + (double) antiguedadMes/6;
                Double brutoMensual = salarioBaseMes+complementosMes+antiguedadMes+prorrateoExtra;
                Double calculoBaseEmpTrab = brutoMensual;
                if(!prorrateo){
                    brutoMensual-=prorrateoExtra;
                    prorrateoExtra=0.0;
                }


                //CALCULO DEDUCCIONES EMPLEADO
                Double ssocial = calculoBaseEmpTrab * listaCuotas.get(0);
                Double desempleo = calculoBaseEmpTrab * listaCuotas.get(1);
                Double formacion = calculoBaseEmpTrab * listaCuotas.get(2);
                Double irpf = brutoMensual * irpfRetencion;
                Double totaldeducciones = ssocial + desempleo + formacion + irpf;

                Double liquidomensual = brutoMensual - totaldeducciones;

                //CALCULO GASTOS EMPRESARIO
                Double ssocialEmp = calculoBaseEmpTrab * listaCuotas.get(4);
                Double desempleoEmp = calculoBaseEmpTrab * listaCuotas.get(6);
                Double fogasa = calculoBaseEmpTrab * listaCuotas.get(5);
                Double formacionEmp = calculoBaseEmpTrab * listaCuotas.get(7);
                Double accidentes = calculoBaseEmpTrab * listaCuotas.get(3);
                Double totalempresario = ssocialEmp + desempleoEmp + fogasa + formacionEmp + accidentes;

                Double costeRealEmp = brutoMensual + totalempresario;

                //IMPRESION
                System.out.println("----------------------");
                String empresaInfo = "Empresa: "+data.get(key).get(4)+", CIF: "+ data.get(key).get(3);
                System.out.println(empresaInfo);

                String trabajadorInfo = data.get(key).get(0)+" "+data.get(key).get(1)+" "+data.get(key).get(2)+", DNI: "+ key;
                System.out.println(trabajadorInfo);

                String trabajadorInfo2 = "IBAN: "+data.get(key).get(10)+", Categoria: "+data.get(key).get(6)+ ", Fecha de alta: "+data.get(key).get(5);
                trabajadorInfo2+=", Bruto Anual: "+prec(brutoAnual);
                System.out.println(trabajadorInfo2);

                System.out.println("Nomina: "+getMonthName(mesCalculo)+" de "+añoCalculo);

                String importes = "IMPORTES DEL TRABAJADOR: \n  Salario base mes: "+prec(salarioBaseMes)+", prorrateo mes: "+prec(prorrateoExtra);
                importes+=", complemento mes: "+prec(complementosMes)+", antiguedad mes: "+antiguedadMes;
                System.out.println(importes);

                String descuentos = "DESCUENTOS DEL TRABAJADOR: \n";
                descuentos+= "  Seguridad Social: "+prec(listaCuotas.get(0)*100.0)+"% de "+prec(calculoBaseEmpTrab)+": "+prec(ssocial)+"\n";
                descuentos+="  Desempleo: "+prec(listaCuotas.get(1)*100.0)+"% de "+prec(calculoBaseEmpTrab)+": "+prec(desempleo)+"\n";
                descuentos+= "  Cuota de formacion: "+prec(listaCuotas.get(2)*100.0)+"% de "+prec(calculoBaseEmpTrab)+": "+prec(formacion)+"\n";
                descuentos+= "  IRPF: "+prec(irpfRetencion*100.0)+"% de "+prec(brutoMensual)+": "+prec(irpf);
                System.out.println(descuentos);

                String ingresos = "TOTAL INGRESOS Y DEDUCCIONES DEL TRABAJADOR: \n";
                ingresos+= "  Devengos: "+prec(brutoMensual)+", Deducciones: "+prec(totaldeducciones)+", Liquido mensual: "+prec(liquidomensual);
                System.out.println(ingresos);

                String empresario="PAGOS DEL EMPRESARIO: \n  Base del calculo sobre el empresario: "+prec(calculoBaseEmpTrab)+"\n";
                empresario+="  Seguridad Social: "+prec(listaCuotas.get(4)*100.0)+"%: "+prec(ssocialEmp)+"\n";
                empresario+="  Desempleo: "+prec(listaCuotas.get(6)*100.0)+"%: "+prec(desempleoEmp)+"\n";
                empresario+= "  Formacion: "+prec(listaCuotas.get(7)*100.0)+"%: "+prec(formacionEmp)+"\n";
                empresario+= "  FOGASA: "+prec(listaCuotas.get(5)*100.0)+"%: "+prec(fogasa)+"\n";
                empresario+= "  Accidentes de trabajo: "+prec(listaCuotas.get(3)*100.0)+"%: "+prec(accidentes)+"\n";
                empresario+= "  Total empresario: "+prec(totalempresario)+"\n";
                empresario+= "  COSTE TOTAL DEL TRABAJADOR: "+prec(costeRealEmp);
                System.out.println(empresario);
                //FIN IMPRESION

                //CALCULO DE EXTRAS
                if(!prorrateo && (mesCalculo==6 || mesCalculo==12)){
                    //COMPROBAMOS PORCENTAJES DE EXTRAS POR SI HAY UN EMPLEADO RECIÉN CONTRATADO
                    Double porcentajeExtra=1.0;
                    if(daysDiff<365) {
                        Double[] extras = getExtrasRecienContratado(dateAltaString, mainDateString);
                        if(mesCalculo==12) porcentajeExtra = extras[1];
                        if(mesCalculo==6) porcentajeExtra = extras[2];
                    }

                    Double brutoMensualExtra = brutoMensual * porcentajeExtra;
                    Double salarioBaseExtra = salarioBaseMes * porcentajeExtra;
                    Double complementoExtra = complementosMes * porcentajeExtra;

                    Double irpfExtra = brutoMensualExtra * irpfRetencion; Double totaldeduccionExtra = irpfExtra;
                    Double liquidoextra = brutoMensualExtra-totaldeduccionExtra;
                    Double costeempresarioextra = brutoMensualExtra;

                    //IMPRESION DE EXTRA
                    System.out.println("----------------------");
                    System.out.println(empresaInfo);
                    System.out.println(trabajadorInfo);
                    System.out.println(trabajadorInfo2);
                    System.out.println("Nomina: Extra de "+getMonthName(mesCalculo)+" de "+añoCalculo);

                    String importesextra = "IMPORTES DEL TRABAJADOR: \n  Salario base mes: "+prec(salarioBaseExtra)+", prorrateo mes: "+prec(prorrateoExtra);
                    importesextra+=", complemento mes: "+prec(complementoExtra)+", antiguedad mes: "+antiguedadMes;
                    System.out.println(importesextra);

                    String descuentosextra = "DESCUENTOS DEL TRABAJADOR: \n";
                    descuentosextra+= "  Seguridad Social: "+prec(listaCuotas.get(0)*100.0)+"% de 0.0: 0.0 \n";
                    descuentosextra+="  Desempleo: "+prec(listaCuotas.get(1)*100.0)+"% de 0.0: 0.0 \n";
                    descuentosextra+= "  Cuota de formacion: "+prec(listaCuotas.get(2)*100.0)+"% de 0.0: 0.0 \n";
                    descuentosextra+= "  IRPF: "+prec(irpfRetencion*100.0)+"% de "+prec(brutoMensualExtra)+": "+prec(irpfExtra);
                    System.out.println(descuentosextra);

                    String ingresosextra = "TOTAL INGRESOS Y DEDUCCIONES DEL TRABAJADOR: \n";
                    ingresosextra+= "  Devengos: "+prec(brutoMensualExtra)+", Deducciones: "+prec(totaldeduccionExtra)+", Liquido mensual: "+prec(liquidoextra);
                    System.out.println(ingresosextra);

                    String empresarioextra="PAGOS DEL EMPRESARIO: \n  Base del calculo sobre el empresario: 0.0\n";
                    empresarioextra+="  Seguridad Social: "+prec(listaCuotas.get(4)*100.0)+"%: 0.0\n";
                    empresarioextra+="  Desempleo: "+prec(listaCuotas.get(6)*100.0)+"%: 0.0\n";
                    empresarioextra+= "  Formacion: "+prec(listaCuotas.get(7)*100.0)+"%: 0.0\n";
                    empresarioextra+= "  FOGASA: "+prec(listaCuotas.get(5)*100.0)+"%: 0.0\n";
                    empresarioextra+= "  Accidentes de trabajo: "+prec(listaCuotas.get(3)*100.0)+"%: 0.0\n";
                    empresarioextra+= "  Total empresario: 0.0\n";
                    empresarioextra+= "  COSTE TOTAL DEL TRABAJADOR: "+prec(costeempresarioextra);
                    System.out.println(empresarioextra);
                    //FIN IMPRESION
                }
            }
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
        int añoActual = Integer.parseInt(mainDate.substring(6));
        int añoAlta = Integer.parseInt(dateAlta.substring(6));

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
        if(bruto<=lista.get(0).getBrutoAnual()) return lista.get(0).getRetencion();

        for(int i=0; i<(lista.size()-1); i++){
            if(bruto>lista.get(i).getBrutoAnual() && bruto<=lista.get(i+1).getBrutoAnual()){
                return lista.get(i+1).getRetencion();
            }
        }

        return lista.get(lista.size()-1).getRetencion();
    }

    public static void PDFgenerator(String key, HashMap<String,List<String>> data) throws IOException {
        File yourFile = new File(".\\resources\\prueba.pdf");
        yourFile.createNewFile();
        PdfWriter writer = new PdfWriter(new FileOutputStream(yourFile));
        PdfDocument pdfDoc = new PdfDocument(writer);
        Document doc = new Document(pdfDoc, PageSize.LETTER);

        Paragraph empty = new Paragraph("");
        Table tabla1 = new Table(2);
        tabla1.setWidth(500);

        //String dateAltaString = data.get(key).get(5);

        Paragraph nom = new Paragraph("NOMBRE");
        Paragraph cif = new Paragraph("CIF: ");

        Paragraph dir1 = new Paragraph("Avenida de la facultad - 6");
        Paragraph dir2 = new Paragraph("24001 León");

        Cell cell1 = new Cell();
        cell1.setBorder(new SolidBorder(1));
        cell1.setWidth(250);
        cell1.setTextAlignment(TextAlignment.CENTER);

        cell1.add(nom);
        cell1.add(cif);
        cell1.add(dir1);
        cell1.add(dir2);
        tabla1.addCell(cell1);

        Cell cell2 = new Cell();
        cell2.setBorder(Border.NO_BORDER);
        cell2.setPadding(10);
        cell2.setTextAlignment(TextAlignment.RIGHT);
        cell2.add(new Paragraph("IBAN: "));
        cell2.add(new Paragraph("Bruto anual: "));
        cell2.add(new Paragraph("Categoría: "));
        cell2.add(new Paragraph("Fecha de alta: "));
        tabla1.addCell(cell2);

        Table tabla2 = new Table(2);

        tabla2.setWidth(500);
        Image img = new Image(ImageDataFactory.create(".\\resources\\logo\\logo.png"));
        img.setBorder(Border.NO_BORDER);
        img.setPadding(10);

        Cell cell3 = new Cell();
        cell3.add(img);
        cell3.setBorder(Border.NO_BORDER);
        cell3.setPaddingLeft(23);
        cell3.setPaddingTop(20);

        cell3.setWidth(250);
        tabla2.addCell(cell3);

        doc.add(tabla1);
        doc.add(tabla2);

        doc.close();
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
            Double retencion = sheet.getRow(r).getCell(1).getNumericCellValue()/100;
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
            double cuota =  (double) sheet.getRow(r).getCell(1).getNumericCellValue()/100;
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

    public static String getMonthName(int m){
        switch (m){
            case 1: return "Enero";
            case 2: return "Febrero";
            case 3: return "Marzo";
            case 4: return "Abril";
            case 5: return "Mayo";
            case 6: return "Junio";
            case 7: return "Julio";
            case 8: return "Agosto";
            case 9: return "Septiembre";
            case 10: return "Octubre";
            case 11: return "Noviembre";
            case 12: return "Diciembre";
        }

        return "";
    }
}

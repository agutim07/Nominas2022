package excelManager;

import map.*;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.ParseException;
import java.util.*;

import java.text.SimpleDateFormat;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class NominasGenerator {
    public static void generarNominas(List<Trabajadorbbdd> data, List<Categorias> listaCategorias, List<Nomina> listaNominas, XSSFWorkbook workbook) throws ParseException, IOException {
        //CONSEGUIMOS LOS DATOS DE LAS HOJAS 1-4
        getHoja3(listaCategorias,workbook); //la hoja 1 se  guarda en listaCategorias
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

        for(int i=0; i<data.size(); i++){
            Trabajadorbbdd iter = data.get(i);
            String key = iter.getNifnie();

            //CHECKEAMOS LA FECHA DE ALTA EN LA EMPRESA CON LA INTRODUCIDA
            Date dateAlta = iter.getFechaAlta();
            SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
            String dateAltaString = dateFormat.format(dateAlta);

            TimeUnit time = TimeUnit.DAYS;
            long daysDiff = time.convert(mainDate.getTime() - dateAlta.getTime(), TimeUnit.MILLISECONDS);

            //SI DAYSDIFF<28 NO SE CALCULA LA NOMINA
            if(daysDiff>=28){
                int añoCalculo = Integer.parseInt(mainDateString.substring(6));
                int trienio = (añoCalculo-Integer.parseInt(dateAltaString.substring(6)))/3;
                int mesCalculo = Integer.valueOf(mainDateString.substring(3,5));

                String pro = iter.getProrrateo(); boolean prorrateo=false;
                if(pro.equals("SI")) prorrateo=true;

                String categoria = iter.getCategorias().getNombreCategoria();
                Double[] salariocomp = getSalarioyComplementos(categoria,listaCategorias);
                Double salarioBase = salariocomp[0]; Double complementos = salariocomp[1];
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

                //IMPRESION Y CREACCION DE OBJETO NOMINA
                Nomina nom = new Nomina();
                nom.setTrabajadorbbdd(data.get(i));
                nom.setNumeroTrienios(trienio);

                System.out.println("----------------------");
                String empresaInfo = "Empresa: "+iter.getEmpresas().getNombre()+", CIF: "+ iter.getEmpresas().getCif();
                System.out.println(empresaInfo);

                String trabajadorInfo = iter.getNombre()+" "+iter.getApellido1()+" "+iter.getApellido2()+", DNI: "+ key;
                System.out.println(trabajadorInfo);

                String trabajadorInfo2 = "IBAN: "+iter.getIban()+", Categoria: "+iter.getCategorias().getNombreCategoria()+ ", Fecha de alta: "+dateAltaString;
                trabajadorInfo2+=", Bruto Anual: "+prec(brutoAnual);
                System.out.println(trabajadorInfo2);
                nom.setBrutoAnual(prec(brutoAnual));

                System.out.println("Nomina: "+getMonthName(mesCalculo)+" de "+añoCalculo);
                nom.setMes(mesCalculo); nom.setAnio(añoCalculo);

                String importes = "IMPORTES DEL TRABAJADOR: \n  Salario base mes: "+prec(salarioBaseMes)+", prorrateo mes: "+prec(prorrateoExtra);
                importes+=", complemento mes: "+prec(complementosMes)+", antiguedad mes: "+antiguedadMes;
                System.out.println(importes);
                nom.setImporteTrienios((double) antiguedadMes); nom.setImporteSalarioMes(prec(salarioBaseMes));
                nom.setImporteComplementoMes(prec(complementosMes)); nom.setValorProrrateo(prec(prorrateoExtra));

                String descuentos = "DESCUENTOS DEL TRABAJADOR: \n";
                descuentos+= "  Seguridad Social: "+prec(listaCuotas.get(0)*100.0)+"% de "+prec(calculoBaseEmpTrab)+": "+prec(ssocial)+"\n";
                nom.setSeguridadSocialTrabajador(prec(listaCuotas.get(0)*100.0)); nom.setImporteSeguridadSocialTrabajador(prec(ssocial));
                descuentos+="  Desempleo: "+prec(listaCuotas.get(1)*100.0)+"% de "+prec(calculoBaseEmpTrab)+": "+prec(desempleo)+"\n";
                nom.setDesempleoTrabajador(prec(listaCuotas.get(1)*100.0)); nom.setImporteDesempleoTrabajador(prec(desempleo));
                descuentos+= "  Cuota de formacion: "+prec(listaCuotas.get(2)*100.0)+"% de "+prec(calculoBaseEmpTrab)+": "+prec(formacion)+"\n";
                nom.setFormacionTrabajador(prec(listaCuotas.get(2)*100.0)); nom.setImporteFormacionTrabajador(prec(formacion));
                descuentos+= "  IRPF: "+prec(irpfRetencion*100.0)+"% de "+prec(brutoMensual)+": "+prec(irpf);
                nom.setIrpf(prec(irpfRetencion*100.0)); nom.setImporteIrpf(prec(irpf));
                System.out.println(descuentos);

                String ingresos = "TOTAL INGRESOS Y DEDUCCIONES DEL TRABAJADOR: \n";
                ingresos+= "  Devengos: "+prec(brutoMensual)+", Deducciones: "+prec(totaldeducciones)+", Liquido mensual: "+prec(liquidomensual);
                nom.setBrutoNomina(prec(brutoMensual)); nom.setLiquidoNomina(prec(liquidomensual));
                System.out.println(ingresos);

                String empresario="PAGOS DEL EMPRESARIO: \n  Base del calculo sobre el empresario: "+prec(calculoBaseEmpTrab)+"\n";
                nom.setBaseEmpresario(prec(calculoBaseEmpTrab));
                empresario+="  Seguridad Social: "+prec(listaCuotas.get(4)*100.0)+"%: "+prec(ssocialEmp)+"\n";
                nom.setSeguridadSocialEmpresario(prec(listaCuotas.get(4)*100.0)); nom.setImporteSeguridadSocialEmpresario(prec(ssocialEmp));
                empresario+="  Desempleo: "+prec(listaCuotas.get(6)*100.0)+"%: "+prec(desempleoEmp)+"\n";
                nom.setDesempleoEmpresario(prec(listaCuotas.get(6)*100.0)); nom.setImporteDesempleoEmpresario(prec(desempleoEmp));
                empresario+= "  Formacion: "+prec(listaCuotas.get(7)*100.0)+"%: "+prec(formacionEmp)+"\n";
                nom.setFormacionEmpresario(prec(listaCuotas.get(7)*100.0)); nom.setImporteFormacionEmpresario(prec(formacionEmp));
                empresario+= "  FOGASA: "+prec(listaCuotas.get(5)*100.0)+"%: "+prec(fogasa)+"\n";
                nom.setFogasaempresario(prec(listaCuotas.get(5)*100.0)); nom.setImporteFogasaempresario(prec(fogasa));
                empresario+= "  Accidentes de trabajo: "+prec(listaCuotas.get(3)*100.0)+"%: "+prec(accidentes)+"\n";
                nom.setAccidentesTrabajoEmpresario(prec(listaCuotas.get(3)*100.0)); nom.setImporteAccidentesTrabajoEmpresario(prec(accidentes));
                empresario+= "  Total empresario: "+prec(totalempresario)+"\n";
                empresario+= "  COSTE TOTAL DEL TRABAJADOR: "+prec(costeRealEmp);
                nom.setCosteTotalEmpresario(prec(costeRealEmp));
                System.out.println(empresario);

                listaNominas.add(nom);
                Set nominasCollecion = data.get(i).getNominas();
                nominasCollecion.add(nom);
                data.get(i).setNominas(nominasCollecion);
                //FIN IMPRESION Y CREACION

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

                    //IMPRESION DE EXTRA Y CREACCION DE OBJETO NOMINA
                    Nomina nomExtra = new Nomina();
                    nomExtra.setTrabajadorbbdd(data.get(i));
                    nomExtra.setNumeroTrienios(trienio);
                    nomExtra.setMes(mesCalculo); nomExtra.setAnio(añoCalculo);
                    nomExtra.setBrutoAnual(prec(brutoAnual));

                    System.out.println("----------------------");
                    System.out.println(empresaInfo);
                    System.out.println(trabajadorInfo);
                    System.out.println(trabajadorInfo2);
                    System.out.println("Nomina: Extra de "+getMonthName(mesCalculo)+" de "+añoCalculo);

                    String importesextra = "IMPORTES DEL TRABAJADOR: \n  Salario base mes: "+prec(salarioBaseExtra)+", prorrateo mes: "+prec(prorrateoExtra);
                    importesextra+=", complemento mes: "+prec(complementoExtra)+", antiguedad mes: "+antiguedadMes;
                    System.out.println(importesextra);
                    nomExtra.setImporteSalarioMes(prec(salarioBaseExtra)); nomExtra.setImporteTrienios((double) antiguedadMes);
                    nomExtra.setImporteComplementoMes(prec(complementoExtra)); nomExtra.setValorProrrateo(prec(prorrateoExtra));

                    String descuentosextra = "DESCUENTOS DEL TRABAJADOR: \n";
                    descuentosextra+= "  Seguridad Social: "+prec(listaCuotas.get(0)*100.0)+"% de 0.0: 0.0 \n";
                    nomExtra.setSeguridadSocialTrabajador(nom.getSeguridadSocialTrabajador()); nomExtra.setImporteSeguridadSocialTrabajador(0.0);
                    descuentosextra+="  Desempleo: "+prec(listaCuotas.get(1)*100.0)+"% de 0.0: 0.0 \n";
                    nomExtra.setImporteDesempleoTrabajador(0.0); nomExtra.setDesempleoTrabajador(nom.getDesempleoTrabajador());
                    descuentosextra+= "  Cuota de formacion: "+prec(listaCuotas.get(2)*100.0)+"% de 0.0: 0.0 \n";
                    nomExtra.setImporteFormacionTrabajador(0.0); nomExtra.setFormacionTrabajador(nom.getFormacionTrabajador());
                    descuentosextra+= "  IRPF: "+prec(irpfRetencion*100.0)+"% de "+prec(brutoMensualExtra)+": "+prec(irpfExtra);
                    nomExtra.setImporteIrpf(prec(irpfExtra)); nomExtra.setIrpf(nom.getIrpf());
                    System.out.println(descuentosextra);

                    String ingresosextra = "TOTAL INGRESOS Y DEDUCCIONES DEL TRABAJADOR: \n";
                    ingresosextra+= "  Devengos: "+prec(brutoMensualExtra)+", Deducciones: "+prec(totaldeduccionExtra)+", Liquido mensual: "+prec(liquidoextra);
                    System.out.println(ingresosextra);
                    nomExtra.setBrutoNomina(prec(brutoMensualExtra)); nomExtra.setLiquidoNomina(prec(liquidoextra));

                    String empresarioextra="PAGOS DEL EMPRESARIO: \n  Base del calculo sobre el empresario: 0.0\n";
                    nomExtra.setBaseEmpresario(0.0);
                    empresarioextra+="  Seguridad Social: "+prec(listaCuotas.get(4)*100.0)+"%: 0.0\n";
                    nomExtra.setImporteSeguridadSocialEmpresario(0.0);  nomExtra.setSeguridadSocialEmpresario(nom.getSeguridadSocialEmpresario());
                    empresarioextra+="  Desempleo: "+prec(listaCuotas.get(6)*100.0)+"%: 0.0\n";
                    nomExtra.setImporteDesempleoEmpresario(0.0); nomExtra.setDesempleoEmpresario(nom.getDesempleoEmpresario());
                    empresarioextra+= "  Formacion: "+prec(listaCuotas.get(7)*100.0)+"%: 0.0\n";
                    nomExtra.setImporteFormacionEmpresario(0.0); nomExtra.setFormacionEmpresario(nom.getFormacionEmpresario());
                    empresarioextra+= "  FOGASA: "+prec(listaCuotas.get(5)*100.0)+"%: 0.0\n";
                    nomExtra.setImporteFogasaempresario(0.0); nomExtra.setFogasaempresario(nom.getFogasaempresario());
                    empresarioextra+= "  Accidentes de trabajo: "+prec(listaCuotas.get(3)*100.0)+"%: 0.0\n";
                    nomExtra.setImporteAccidentesTrabajoEmpresario(0.0); nomExtra.setAccidentesTrabajoEmpresario(nom.getAccidentesTrabajoEmpresario());
                    empresarioextra+= "  Total empresario: 0.0\n";
                    empresarioextra+= "  COSTE TOTAL DEL TRABAJADOR: "+prec(costeempresarioextra);
                    nomExtra.setCosteTotalEmpresario(prec(costeempresarioextra));
                    System.out.println(empresarioextra);

                    listaNominas.add(nomExtra);
                    nominasCollecion.add(nomExtra);
                    data.get(i).setNominas(nominasCollecion);
                    //FIN IMPRESION Y CREACION
                }
            }
        }
    }

    public static Double[] getSalarioyComplementos(String cat, List<Categorias> list){
        for(int i=0; i< list.size(); i++){
            if(cat.equals(list.get(i).getNombreCategoria())){
                return new Double[]{list.get(i).getSalarioBaseCategoria(),list.get(i).getComplementoCategoria()};
            }
        }

        return new Double[]{0.0,0.0};
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

    public static void getHoja3(List<Categorias> categorias, XSSFWorkbook workbook){
        //INTRODUCIMOS HOJA 3: CATEGORIA, COMPLEMENTO Y SALARIO BASE
        XSSFSheet sheet = workbook.getSheet("Hoja3");
        int rows = sheet.getLastRowNum();

        for(int r=1; r<=rows; r++) {
            String cat = sheet.getRow(r).getCell(0).getStringCellValue();
            for(int i=0; i< categorias.size(); i++){
                if(categorias.get(i).getNombreCategoria().equals(cat)){
                    categorias.get(i).setComplementoCategoria(sheet.getRow(r).getCell(1).getNumericCellValue());
                    categorias.get(i).setSalarioBaseCategoria(sheet.getRow(r).getCell(2).getNumericCellValue());
                    break;
                }
            }
        }
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
            double cuota =  sheet.getRow(r).getCell(1).getNumericCellValue()/100;
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

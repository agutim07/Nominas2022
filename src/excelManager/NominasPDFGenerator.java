package excelManager;

import com.itextpdf.io.font.constants.FontStyles;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.pdf.canvas.draw.SolidLine;
import com.itextpdf.layout.element.*;
import com.itextpdf.layout.properties.*;
import map.*;
import java.io.File;
import java.io.FileOutputStream;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.kernel.colors.ColorConstants;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;

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
import com.itextpdf.io.font.constants.StandardFonts;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.List;

public class NominasPDFGenerator {
    public static void generarPDFs(List<Nomina> listaNominas) throws IOException {
        for(int i=0; i<1; i++){
            PDFgenerator(listaNominas.get(i));
        }
    }

    private static void PDFgenerator(Nomina n) throws IOException {
        Trabajadorbbdd trab = n.getTrabajadorbbdd();
        Empresas emp = trab.getEmpresas();
        boolean isExtra = false;
        if(n.getBaseEmpresario()==0.0) isExtra = true;

        String nombreFile = trab.getNifnie()+trab.getNombre()+trab.getApellido1()+trab.getApellido2()+getMonthName(n.getMes())+n.getAnio();
        if(isExtra) nombreFile += "EXTRA";
        File yourFile = new File(".\\resources\\nominas\\"+nombreFile+".pdf");
        yourFile.createNewFile();
        PdfWriter writer = new PdfWriter(new FileOutputStream(yourFile));
        PdfDocument pdfDoc = new PdfDocument(writer);
        Document doc = new Document(pdfDoc, PageSize.LETTER);

        PdfFont bold = PdfFontFactory.createFont(StandardFonts.COURIER_BOLD);
        PdfFont bold_ob = PdfFontFactory.createFont(StandardFonts.COURIER_BOLDOBLIQUE);
        PdfFont font = PdfFontFactory.createFont(StandardFonts.COURIER);

        //TABLA 1
        Table tabla1 = new Table(2);
        tabla1.setWidth(500);

        Cell cell1 = new Cell();
        cell1.setBorder(new SolidBorder(1));
        cell1.setWidth(250);
        cell1.setTextAlignment(TextAlignment.CENTER);
        cell1.setVerticalAlignment(VerticalAlignment.MIDDLE);

        cell1.add(new Paragraph("EMPRESA").setFont(bold));
        cell1.add(new Paragraph(emp.getNombre()).setFont(font));
        cell1.add(new Paragraph("CIF: "+emp.getCif()).setFont(font));

        tabla1.addCell(cell1);

        Cell cell2 = new Cell();
        cell2.setBorder(Border.NO_BORDER);
        cell2.setPadding(10);
        cell2.setTextAlignment(TextAlignment.RIGHT);

        cell2.add(new Paragraph("IBAN: "+trab.getIban())).setFont(font);
        cell2.add(new Paragraph("Bruto anual: "+n.getBrutoAnual())).setFont(font);
        cell2.add(new Paragraph("Categoría: "+trab.getCategorias().getNombreCategoria())).setFont(font);
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        String fechaAlta = dateFormat.format(trab.getFechaAlta());
        cell2.add(new Paragraph("Fecha de alta: "+fechaAlta)).setFont(font);

        tabla1.addCell(cell2);

        //TABLA 2
        Table tabla2 = new Table(2);
        tabla2.setBorderCollapse(BorderCollapsePropertyValue.SEPARATE);
        tabla2.setVerticalBorderSpacing(10);
        tabla2.setWidth(500);

        Image img = new Image(ImageDataFactory.create(".\\resources\\logo\\logo.png"));
        img.setBorder(Border.NO_BORDER);
        img.setPadding(10);
        img.scaleAbsolute((float) (80f*3.03), 80f);

        Cell cell3 = new Cell();
        cell3.add(img);
        cell3.setBorder(Border.NO_BORDER);
        cell3.setPaddingLeft(23);
        cell3.setPaddingTop(20);

        tabla2.addCell(cell3);

        Cell cell4 = new Cell();
        cell4.setBorder(new SolidBorder(1));
        cell4.setWidth(200);
        cell4.setPadding(10);
        cell4.setTextAlignment(TextAlignment.RIGHT);
        cell4.setVerticalAlignment(VerticalAlignment.TOP);

        cell4.add(new Paragraph("Destinatario:").setFont(bold).setTextAlignment(TextAlignment.LEFT).setMultipliedLeading(3.0f));
        cell4.add(new Paragraph(trab.getNombre()+" "+trab.getApellido1()+" "+trab.getApellido2()).setFont(font));
        cell4.add(new Paragraph("DNI: "+trab.getNifnie()).setFont(font));

        tabla2.addCell(cell4);

        //TABLA 3
        Table tabla3 = new Table(1);
        tabla2.setBorderCollapse(BorderCollapsePropertyValue.SEPARATE);
        tabla2.setVerticalBorderSpacing(20);
        tabla2.setWidth(500);

        Cell cell5 = new Cell();
        cell5.setBorder(Border.NO_BORDER);
        cell5.setWidth(500);
        cell5.setPadding(10);
        cell5.setTextAlignment(TextAlignment.CENTER);

        String titulo = "Nómina: "; if(isExtra) titulo+="Extra de ";
        titulo+=getMonthName(n.getMes())+" de "+n.getAnio();
        cell5.add(new Paragraph(titulo).setFont(bold_ob));
        tabla3.addCell(cell5);

        //TABLA 4
        Table tabla4 = new Table(1);
        tabla4.setWidth(500);
        tabla4.setBorder(new SolidBorder(1));

        Cell cell6 = new Cell();
        cell6.setBorder(Border.NO_BORDER);
        cell6.setWidth(500);
        cell6.setPadding(4);
        cell6.setTextAlignment(TextAlignment.CENTER);

        cell6.add(new Paragraph("Trabajador").setFont(bold));
        tabla4.addCell(cell6);

        //TABLA 5
        Table tabla5 = new Table(2);
        tabla5.setWidth(500);
        tabla5.setFontSize(8f);

        Cell cell7 = new Cell();
        cell7.setBorder(new SolidBorder(1));
        cell7.setWidth(250);
        cell7.setTextAlignment(TextAlignment.LEFT);

        Paragraph p1 = new Paragraph("Salario base").setFont(font);
        p1.add(new Tab()); p1.addTabStops(new TabStop(1000, TabAlignment.RIGHT));
        p1.add(""+n.getImporteSalarioMes());
        Paragraph p2 = new Paragraph("Complemento").setFont(font);
        p2.add(new Tab()); p2.addTabStops(new TabStop(1000, TabAlignment.RIGHT));
        p2.add(""+n.getImporteComplementoMes());
        Paragraph p3 = new Paragraph("Prorrateo").setFont(font);
        p3.add(new Tab()); p3.addTabStops(new TabStop(1000, TabAlignment.RIGHT));
        p3.add(""+n.getValorProrrateo());
        Paragraph p4 = new Paragraph("Antiguedad").setFont(font);
        p4.add(new Tab()); p4.addTabStops(new TabStop(1000, TabAlignment.RIGHT));
        p4.add("{"+n.getNumeroTrienios()+" trienio/s} "+n.getImporteTrienios());

        cell7.add(new Paragraph("Importes mensuales").setFont(bold).setTextAlignment(TextAlignment.CENTER));
        cell7.add(p1);
        cell7.add(p2);
        cell7.add(p3);
        cell7.add(p4);
        cell7.add(new Paragraph("TOTAL DEVENGOS: "+n.getBrutoNomina()).setFont(bold).setTextAlignment(TextAlignment.RIGHT).setMultipliedLeading(2.0f).setVerticalAlignment(VerticalAlignment.BOTTOM));

        Cell cell8 = new Cell();
        cell8.setBorder(new SolidBorder(1));
        cell8.setWidth(250);
        cell8.setTextAlignment(TextAlignment.LEFT);

        Paragraph p4_ = new Paragraph("Seguridad Social").setFont(font);
        p4_.add(new Tab()); p4_.addTabStops(new TabStop(1000, TabAlignment.RIGHT));
        p4_.add(n.getSeguridadSocialTrabajador()+"% de "+n.getBaseEmpresario()+": "+n.getImporteSeguridadSocialTrabajador());
        Paragraph p5 = new Paragraph("Desempleo").setFont(font);
        p5.add(new Tab()); p5.addTabStops(new TabStop(1000, TabAlignment.RIGHT));
        p5.add(n.getDesempleoTrabajador()+"% de "+n.getBaseEmpresario()+": "+n.getImporteDesempleoTrabajador());
        Paragraph p6 = new Paragraph("Cuota de formación").setFont(font);
        p6.add(new Tab()); p6.addTabStops(new TabStop(1000, TabAlignment.RIGHT));
        p6.add(n.getFormacionTrabajador()+"% de "+n.getBaseEmpresario()+": "+n.getImporteFormacionTrabajador());
        Paragraph p7 = new Paragraph("IRPF").setFont(font);
        p7.add(new Tab()); p7.addTabStops(new TabStop(1000, TabAlignment.RIGHT));
        p7.add(n.getIrpf()+"% de "+n.getBrutoNomina()+": "+n.getImporteIrpf());

        cell8.add(new Paragraph("Descuentos ").setFont(bold).setTextAlignment(TextAlignment.CENTER));
        cell8.add(p4_);
        cell8.add(p5);
        cell8.add(p6);
        cell8.add(p7);
        cell8.add(new Paragraph("TOTAL DEDUCCIONES ").setFont(bold).setTextAlignment(TextAlignment.RIGHT).setMultipliedLeading(2.0f).setVerticalAlignment(VerticalAlignment.BOTTOM));

        tabla5.addCell(cell7);
        tabla5.addCell(cell8);


        //FIN
        doc.add(tabla1);
        doc.add(tabla2);
        doc.add(tabla3);
        doc.add(tabla4);
        doc.add(tabla5);

        doc.close();

//        SolidLine line = new SolidLine(1f);
//        line.setColor(ColorConstants.RED);
//        LineSeparator ls = new LineSeparator(line);
//        ls.setMarginTop(5);
    }

    private static String getMonthName(int m){
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

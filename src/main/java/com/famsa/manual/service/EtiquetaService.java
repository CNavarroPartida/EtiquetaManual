package com.famsa.manual.service;

import com.famsa.manual.model.Etiqueta;

import com.lowagie.text.DocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.DocumentFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.FileSystemResource;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.thymeleaf.TemplateEngine;
import org.thymeleaf.context.Context;
import org.thymeleaf.templatemode.TemplateMode;
import org.thymeleaf.templateresolver.ClassLoaderTemplateResolver;
import org.xhtmlrenderer.pdf.ITextRenderer;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;

@Service
public class EtiquetaService {

    private static final String EMPTY = "";

    public FileSystemResource generarEtiqueta(MultipartFile file, HttpServletResponse response) throws Exception {

        List<Etiqueta> etiquetas = createEtiquetaListFromExcel(file);
        String templateDirection = "templates/cont4";

        Context context = new Context();
        context.setVariable("listaEtiquetas", etiquetas);

        String html = parseThymeleafTemplate(templateDirection, context);


        // Creamos el archivo virtualmente y lo convertimos en bytes
        String nametmp = "Etiquetas.pdf";
        response.setContentType("application/pdf");
        return new FileSystemResource(createFileResource(html, nametmp));
    }


    private String parseThymeleafTemplate(String templateDirection, Context context) {
        ClassLoaderTemplateResolver templateResolver = new ClassLoaderTemplateResolver();
        templateResolver.setSuffix(".html");
        templateResolver.setTemplateMode(TemplateMode.HTML);
        templateResolver.setCharacterEncoding("UTF-8");

        TemplateEngine templateEngine = new TemplateEngine();
        templateEngine.setTemplateResolver(templateResolver);

        return templateEngine.process(templateDirection, context);
    }

    public static File createFileResource(String message, String name) throws IOException, DocumentFormatException, DocumentException {
        File file = File.createTempFile(name, "PDF");
        ITextRenderer renderer = new ITextRenderer();

        renderer.setDocumentFromString(message);
        renderer.layout();
        renderer.createPDF(new FileOutputStream(file));

        return file;
    }


    private List<Etiqueta> createEtiquetaListFromExcel(MultipartFile file) throws IOException, InvocationTargetException, IllegalAccessException {
        List<Etiqueta> etiquetaList = new ArrayList<>();

        InputStream is = file.getInputStream();
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);

        // Get the first row with field names
        Row headerRow = sheet.getRow(0);

        // Create a map to store the setter methods for each field
        Map<String, Method> setterMap = new HashMap<>();
        for (Method method : Etiqueta.class.getMethods()) {
            if (method.getName().startsWith("set")) {
                setterMap.put(method.getName().substring(3), method);
            }
        }

        // Loop through the rest of the rows to get the data
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row currentRow = sheet.getRow(i);
            Etiqueta etiqueta = new Etiqueta();

            // Loop through all the cells in the row
            for (int j = 0; j < currentRow.getLastCellNum(); j++) {
                Cell currentCell = currentRow.getCell(j) ;
                String fieldName = headerRow.getCell(j).getStringCellValue();

                if (fieldName == null || fieldName.isEmpty()) {
                    break;
                }

                String fieldNameOutput = fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                Method setterMethod = setterMap.get(fieldNameOutput);

                // Check if a setter method exists for the current field
                if (setterMethod != null) {
                    Object value = null;
                    if (currentCell == null) {
                        value = "";
                    } else {
                        if (currentCell.getCellType() == CellType.STRING) {
                            value = currentCell.getStringCellValue();
                        } else if (currentCell.getCellType() == CellType.NUMERIC) {
                            double numericValue = currentCell.getNumericCellValue();
                            value = String.valueOf(numericValue);
                        }
                    }
                    // Invoke the setter method with the string value
                    setterMethod.invoke(etiqueta, value);
                }
            }

            etiquetaList.add(etiqueta);
        }

        return etiquetaList;
    }



}

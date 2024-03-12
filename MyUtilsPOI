/* LAN0 - 2024-03-12  - Clase Java con utilidades para manejar Ficheros xlsx */

import beans.XlsSheet;
import com.google.gson.JsonArray;
import interfaces.InterfaceExcel;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

public class MyUtilsPOI {

    private static org.apache.log4j.Logger log = org.apache.log4j.Logger.getLogger(MyUtilsPOI.class);
    public static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    public static InterfaceExcel getExcelSheet(ByteArrayInputStream bip, String fileType, String sheetName) {
        InterfaceExcel excel = null;

        if (fileType.equalsIgnoreCase("xls")) {
            excel = new XlsSheet();
            excel.init(bip, sheetName);
        } else {
            excel = new beans.XlsxSheet();
            excel.init(bip, sheetName);

        }

        return excel;
    }

    public static String[] getColNamesRow(Row row) {

        String[] resultado = new String[row.getLastCellNum()];

        int nCellNumber = 0;
        for (nCellNumber = 0; nCellNumber < row.getLastCellNum(); nCellNumber++) {
            Cell celda = row.getCell(nCellNumber);
            if (celda != null) {
                resultado[nCellNumber] = celda.getStringCellValue();
            } else {
                resultado[nCellNumber] = "";
            }
        }

        return resultado;
    }

    public static HashMap<String, Object> getDataMapFromRow(Row row, LinkedHashMap<String, String> mapping, String[] columns) {

        int nCellNumber = 0;

        HashMap<String, Object> resultado = new HashMap<String, Object>();

        for (nCellNumber = 0; nCellNumber < row.getLastCellNum(); nCellNumber++) {

            // conseguimos el nombre de la row
            String columnName = columns[nCellNumber];

            // si el nombre de la col es de los campos mapeados entonces lo añadimos al DataMap
            if (mapping.containsKey(columnName)) {
                Cell celda = row.getCell(nCellNumber);
                String db2FieldName = mapping.get(columnName);

                resultado.put(db2FieldName, getCellValue(celda));
            }
        }

        return resultado;

    }

    public static String getValueFromRowWithColumnName(Row row, String myColumnName, String[] myColsNamesXLS) {

        int nCellNumber = 0;

        String resultado = "";

        for (nCellNumber = 0; nCellNumber < myColsNamesXLS.length; nCellNumber++) {

            if (myColumnName.contentEquals(myColsNamesXLS[nCellNumber])) {

                Cell celda = row.getCell(nCellNumber);

                if (null != celda) {

                    resultado = getCellValue(celda);

                } else {

                    resultado = "";
                }
            }

        }
        return resultado;

    }

    public static String getCellValue(Cell celda) {

        String cellValue = "";
        switch (celda.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(celda)) {
                    cellValue = sdf.format(celda.getDateCellValue());
                } else {
                    cellValue = Double.toString(celda.getNumericCellValue());

                    double dphi = celda.getNumericCellValue();
                    if ((dphi - (int) dphi) * 1000 == 0) {
                        cellValue = (int) dphi + "";
                    }
                }
                // cellValue = Double.toString(celda.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_STRING:
                cellValue = celda.getStringCellValue();
                break;
            case Cell.CELL_TYPE_FORMULA:
                // cellValue = celda.getCellFormula();
                switch (celda.getCachedFormulaResultType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        if (HSSFDateUtil.isCellDateFormatted(celda)) {
                            // cellValue = celda.getDateCellValue().toString();
                            cellValue = sdf.format(celda.getDateCellValue());
                        } else {
                            cellValue = Double.toString(celda.getNumericCellValue());

                            double dphi = celda.getNumericCellValue();
                            if ((dphi - (int) dphi) * 1000 == 0) {
                                cellValue = (int) dphi + "";
                            }
                        }
                        break;
                    case Cell.CELL_TYPE_STRING:
                        cellValue = celda.getStringCellValue();
                        break;
                }

                break;
            case Cell.CELL_TYPE_BLANK:
                cellValue = "";
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                cellValue = Boolean.toString(celda.getBooleanCellValue());
                break;
            default:
                cellValue = "";
                break;

        }

        return cellValue;

    }

    public static boolean generateExcel(InfoJsonBean datos, String fileName, String plantilla, String celdaInicio) {

        boolean resultado = false;
        // Flag para saber si tenemos que crear cabecera de columnas o no
        boolean cabecera = true;

        // XLSX
        XSSFWorkbook workbook = null;
        XSSFSheet mySheet = null;

        if (plantilla.isEmpty()) {

            workbook = new XSSFWorkbook();
            mySheet = workbook.createSheet("hoja1");

        } else {

            try {
                workbook = new XSSFWorkbook(plantilla);
                mySheet = workbook.getSheet("Hoja1");
                cabecera = false;
            } catch (IOException ex) {
                log.error(ex.getMessage());
            }
        }

        JSONArray myColumnas = (JSONArray) datos.datos.get("columnas0");
        ArrayList<String> myListaColumnas = new ArrayList<String>();

        Row row = null;

        if (cabecera) {
            row = mySheet.createRow(0);
        }

        for (int nConta = 0; nConta < myColumnas.size(); nConta++) {
            JSONObject myJsonColumn = (JSONObject) myColumnas.get(nConta);
            if (cabecera) {
                row.createCell(nConta).setCellValue(myJsonColumn.get("fieldName").toString());
            }
            myListaColumnas.add(myJsonColumn.get("fieldName").toString());
        }

        JSONArray myResultado = (JSONArray) datos.datos.get("resultado0");
        int[] indices = null;
        if (!celdaInicio.isEmpty()) {
            indices = excelToIndices(celdaInicio);
        }

        for (int nConta = 0; nConta < myResultado.size(); nConta++) {

            row = mySheet.createRow(nConta + 1 + indices[1]);
            JSONObject myJsonColumn = (JSONObject) myResultado.get(nConta);

            for (Object keyObj : myJsonColumn.keySet()) {

                String key = (String) keyObj;
                String value = myJsonColumn.get(key).toString();
                row.createCell(myListaColumnas.indexOf(key) + indices[0] - 1).setCellValue(value);
            }
        }

        // Guardar el archivo XLSX
        saveXlsx(workbook, fileName);

        return resultado;

    }

    public static XSSFWorkbook generateExcel2(InfoJsonBean datos, XSSFWorkbook workbook, boolean plantilla, String celdaInicio) {

        // Flag para saber si tenemos que crear cabecera de columnas o no
        boolean cabecera = true;
        XSSFSheet mySheet = null;

        if (!plantilla) {

            mySheet = workbook.createSheet("Hoja1");

        } else {

            mySheet = workbook.getSheet("Hoja1");
            cabecera = false;
        }

        JSONArray myColumnas = (JSONArray) datos.datos.get("columnas0");
        ArrayList<String> myListaColumnas = new ArrayList<String>();

        Row row = null;

        if (cabecera) {
            row = mySheet.createRow(0);
        }

        for (int nConta = 0; nConta < myColumnas.size(); nConta++) {
            JSONObject myJsonColumn = (JSONObject) myColumnas.get(nConta);
            if (cabecera) {
                row.createCell(nConta).setCellValue(myJsonColumn.get("fieldName").toString());
            }
            myListaColumnas.add(myJsonColumn.get("fieldName").toString());
        }

        JSONArray myResultado = (JSONArray) datos.datos.get("resultado0");
        int[] indices = null;
        if (!celdaInicio.isEmpty()) {
            indices = excelToIndices(celdaInicio);
        }

        for (int nConta = 0; nConta < myResultado.size(); nConta++) {

            row = mySheet.createRow(nConta + 1 + indices[1]);
            JSONObject myJsonColumn = (JSONObject) myResultado.get(nConta);

            for (Object keyObj : myJsonColumn.keySet()) {

                String key = (String) keyObj;
                String value = myJsonColumn.get(key).toString();
                row.createCell(myListaColumnas.indexOf(key) + indices[0] - 1).setCellValue(value);
            }
        }

        // Guardar el archivo XLSX
        // saveXlsx(workbook, fileName);

        return workbook;

    }

    public static boolean saveXlsx(XSSFWorkbook myWorkBook, String fileName) {

        boolean resultado = false;
        FileOutputStream myFileXls = null;

        try {
            myFileXls = new FileOutputStream(fileName);
            myWorkBook.write(myFileXls);
            myFileXls.close();
            resultado = true;
        } catch (IOException ex) {
            log.error(ex.toString());
        } finally {
            try {
                if (myFileXls != null) {
                    myFileXls.close(); // Cierra el FileOutputStream en el bloque finally
                }
            } catch (IOException ex) {
                log.error("Error al cerrar el FileOutputStream: " + ex.toString());
            }
        }

        return resultado;
    }

    public static int[] excelToIndices(String coordinate) {
        // Divide la coordenada en parte alfabética (columna) y numérica (fila)
        int i = 0;
        while (i < coordinate.length() && Character.isLetter(coordinate.charAt(i))) {
            i++;
        }
        String columnPart = coordinate.substring(0, i);
        String rowPart = coordinate.substring(i);

        // Convertir columna de letras a número (0 basado)
        int column = 0;
        for (int j = 0; j < columnPart.length(); j++) {
            column = column * 26 + (columnPart.charAt(j) - 'A' + 1);
        }
        column -= 1; // Ajusta para base 0

        // Convertir fila a base 0
        int row = Integer.parseInt(rowPart) - 1;

        return new int[]{row, column};
    }

    public static boolean replaceContenido(String fileName, HashMap<String, InfoParamBean> parametros) {

        boolean resultado = false;

        try {
            XSSFWorkbook workbook = new XSSFWorkbook(fileName);

            // Itera a través de todas las hojas del archivo Excel
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);

                // Itera a través de todas las filas de la hoja actual
                for (Row row : sheet) {
                    // Itera a través de todas las celdas de la fila actual
                    for (Cell cell : row) {
                        // Verifica si el contenido de la celda coincide con el valor buscado
                        int cellType = cell.getCellType();

                        // Aqui montaremos el bucle del hashMap de parametros

                        for (Map.Entry<String, InfoParamBean> entry : parametros.entrySet()) {
                            // resultado.put(entry.getValue().getName(), entry.getValue().getValue());
                            String valorABuscar = "<" + entry.getValue().getName() + ">";
                            // String
                            if (cell.getCellType() == 1 && cell.getStringCellValue().equals(valorABuscar)) {
                                // System.out.println("Se encontró el valor en la hoja: " + sheet.getSheetName()
                                //        + ", fila: " + (row.getRowNum() + 1) + ", columna: " + (cell.getColumnIndex() + 1));
                                cell.setCellValue(entry.getValue().getValue());
                            }
                        }

                    }
                }
            }

            saveXlsx(workbook, fileName);

        } catch (IOException e) {
            e.printStackTrace();
        }

        return resultado;
    }

    public static XSSFWorkbook replaceContenido2(XSSFWorkbook workbook, HashMap<String, InfoParamBean> parametros) {

        // Itera a través de todas las hojas del archivo Excel
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            XSSFSheet sheet = workbook.getSheetAt(i);

            // Itera a través de todas las filas de la hoja actual
            for (Row row : sheet) {
                // Itera a través de todas las celdas de la fila actual
                for (Cell cell : row) {
                    // Verifica si el contenido de la celda coincide con el valor buscado
                    int cellType = cell.getCellType();

                    // Aqui montaremos el bucle del hashMap de parametros

                    for (Map.Entry<String, InfoParamBean> entry : parametros.entrySet()) {
                        // resultado.put(entry.getValue().getName(), entry.getValue().getValue());
                        String valorABuscar = "<" + entry.getValue().getName() + ">";
                        // String
                        if (cell.getCellType() == 1 && cell.getStringCellValue().equals(valorABuscar)) {
                            // System.out.println("Se encontró el valor en la hoja: " + sheet.getSheetName()
                            //        + ", fila: " + (row.getRowNum() + 1) + ", columna: " + (cell.getColumnIndex() + 1));
                            cell.setCellValue(entry.getValue().getValue());
                        }
                    }
                }
            }
        }

        // saveXlsx(workbook, fileName);

        return workbook;
    }
}

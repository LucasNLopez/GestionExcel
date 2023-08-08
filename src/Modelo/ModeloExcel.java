package Modelo;

import java.io.*;
import java.util.*;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/**
 *
 * @author lucas
 */
public class ModeloExcel {

    Workbook libro;

    public String Importar(File archivo, JTable tablaD) {
        String response = "No se pudo realizar la importacion";

        DefaultTableModel modeloT = new DefaultTableModel();
        try {
            libro = WorkbookFactory.create(new FileInputStream(archivo));
            Sheet hoja = libro.getSheetAt(0);

            Iterator filaIterator = hoja.rowIterator();
            int indiceFila = -1;
            while (filaIterator.hasNext()) {
                indiceFila++;
                Row fila = (Row) filaIterator.next();
                Iterator columnaIterator = fila.cellIterator();
                Object[] listaColumna = new Object[5];
                int indiceColumna = -1;
                while (columnaIterator.hasNext()) {
                    indiceColumna++;
                    Cell celda = (Cell) columnaIterator.next();
                    if (indiceFila == 0) {
                        modeloT.addColumn(celda.getStringCellValue());
                    } else {
                        if (celda != null) {
                            switch (celda.getCellType()) {
                                case Cell.CELL_TYPE_NUMERIC:
                                    listaColumna[indiceColumna] = (int) Math.round(celda.getNumericCellValue());
                                    break;
                                case Cell.CELL_TYPE_STRING:
                                    listaColumna[indiceColumna] = celda.getStringCellValue();
                                    break;
                                default:
                                    listaColumna[indiceColumna] = celda.getDateCellValue();
                            }
                        }
                    }
                }
                if (indiceFila != 0) {
                    modeloT.addRow(listaColumna);
                }
            }
            response = "Importacion exitosa";
        } catch (Exception e) {
        }

        return response;
    }
    
    public String Exportar (File archivo, JTable tablaD){
        String respuesta= "No se realizo con exito la exportacion";
        int numFila=tablaD.getRowCount(), numColumna=tablaD.getColumnCount();
        if (archivo.getName().endsWith("xlsx")) {
            libro=new XSSFWorkbook();
        }else{
            libro=new HSSFWorkbook();
        }
        
        Sheet hoja = libro.createSheet("Datos");
        try {
            for (int i = -1; i < numFila; i++) {
                Row fila = hoja.createRow(i+1);
                for (int j = 0; j < numColumna; j++) {
                    Cell celda = fila.createCell(i);
                    if (i==-1) {
                        celda.setCellValue(String.valueOf(tablaD.getColumnName(j)));
                    }else{
                        celda.setCellValue(String.valueOf(tablaD.getValueAt(i, i)));
                    }
                    libro.write(new FileOutputStream(archivo));
                }
            }
            respuesta="Exportacion Exitosa";
        } catch (Exception e) {
        }
        return respuesta;
    }
    
}

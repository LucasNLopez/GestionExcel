
package Controlador;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import Vista.VistaExcel;
import Modelo.ModeloExcel;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
/**
@author lucas
 */
public class ControladorExcel implements ActionListener {
    ModeloExcel modeloE= new ModeloExcel();
    VistaExcel vistaE= new VistaExcel();
    JFileChooser selecArchivo= new JFileChooser();
    File archivo;
    int contadorAccion;
    
    public ControladorExcel(VistaExcel vistaE, ModeloExcel modeloE){
        this.modeloE= modeloE;
        this.vistaE=vistaE;
        this.vistaE.btnImportar.addActionListener(this);
        this.vistaE.btnExportar.addActionListener(this);
        
    }
    
    
    
    @Override
    public void actionPerformed(ActionEvent ae) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
    
}

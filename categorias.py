import pandas as pd
import sys

def categorias ():
    print("Asegurese de que el archivo que intenta utilizar se encuentre en la misma dirección que el Script.")
    name=input("Ingrese el nombre del archivo Excel : \n")
    name+=".xlsx"
    try:
    #Lectura del excel
        df=pd.read_excel(name)
    except FileNotFoundError:
        print(f"Error: El archivo {name} no se encuentra en la ruta especificada. Por favor, verifique la ruta y el nombre del archivo.")
        sys.exit()
    except (pd.errors.EmptyDataError):
        print(f"Error: El archivo {name} está vacío o no tiene el formato correcto.")
        sys.exit()
    archivos_creados = 0
    #Recorte 70
    try:
        df_70 = df[df["Producto Desc"].str.contains("70 vl", case=False)]
        df_70=df_70[["Nro Serie"]]
        if not df_70.empty:
            df_70VL = pd.DataFrame({"Nro Etiqueta": df_70["Nro Serie"], "Venta": "", "Categoria": "Trimming 70vl"})
            df_70VL.to_excel("70vl.xlsx", index=False)
            archivos_creados+=1
    #Huesos
        df_huesos=df[df["Producto Desc"].str.startswith("HUESO")]
        df_huesos=df_huesos[["Nro Serie"]]
        if not df_huesos.empty:
            df_huesos=pd.DataFrame({"Nro Etiqueta":df_huesos["Nro Serie"],"Venta":"","Categoria":"Mix Huesos"})
            df_huesos.to_excel("huesos.xlsx",index=False)
            archivos_creados+=1
    #CUADRIL
        df_cuadril= df[df["Producto Desc"].str.startswith("CUADRIL (CAJA)")]
        df_cuadril=df_cuadril[["Nro Serie"]]
        if not df_cuadril.empty:
            df_cuadril= pd.DataFrame({"Nro Etiqueta": df_cuadril["Nro Serie"], "Venta": "", "Categoria": "RUMP"})
            df_cuadril.to_excel("Rump.xlsx", index=False)
            archivos_creados+=1
    #GRASA DE PECHO
        df_grasa_pecho= df[df["Producto Desc"].str.startswith("GRASA DE PECHO (CAJA)")]
        df_grasa_pecho=df_grasa_pecho[["Nro Serie"]]
        if not df_grasa_pecho.empty:
            df_grasa_pecho= pd.DataFrame({"Nro Etiqueta": df_grasa_pecho["Nro Serie"], "Venta": "", "Categoria": "Mix Cuts"})
            df_grasa_pecho.to_excel("Mix Cuts.xlsx", index=False)
            archivos_creados+=1
    #R&L
        valores_a_buscar = ["BIFE ANCHO S/T-+2 kg. (CAJA) ", "BIFE ANCHO S/T-+ 2,5 (CAJA)", "BIFE ANCHO S/T--2 kg. (CAJA)",
                            "BIFE ANGOSTO CON CORDON-+3,5 kg. (CAJA) ","BIFE ANGOSTO CON CORDON-2,5 A 3 (CAJA)",
                            "BIFE ANGOSTO CON CORDON-3 A 3,5 (CAJA)","BIFE ANGOSTO CON CORDON-3 A 3.5 kg. (CAJA) ",
                            "BIFE ANGOSTO CON CORDON--3 kg. (CAJA) ","BIFE ANGOSTO CON CORDON-3,5 A 4 (CAJA)","BIFE ANGOSTO CON CORDON-4 A 5 (CAJA)",
                            "LOMO S/C +5 (CAJA)","LOMO S/C 2 LB (CAJA)","LOMO S/C 2/3 LB (CAJA)",
                            "LOMO S/C 3/4 LB (CAJA)","LOMO S/C 4/5 LB (CAJA)"]
        df_filtrado = pd.DataFrame()  # Crea un DataFrame vacío para almacenar los resultados
        for valor in valores_a_buscar:
            filtro = df["Producto Desc"] == valor
            df_filtrado = pd.concat([df_filtrado, df[filtro]], ignore_index=True)
        #df_filtrado=df_filtrado[["Nro Serie"]]
        #R&L CUPO
        ral_cupo=df_filtrado[pd.isna(df_filtrado["Contramarca"])]
        if not ral_cupo.empty:
            ral_cupo = pd.DataFrame({"Nro Etiqueta": ral_cupo["Nro Serie"], "Venta": "", "Categoria": "R&L"})
            ral_cupo.to_excel("R&L CUPO.xlsx",index=False)
            archivos_creados+=1
        #R&L D/E
        ral_de_filtro=["D/E","T"]
        ral_de=df_filtrado[df_filtrado["Contramarca"].isin(ral_de_filtro)]
        if not ral_de.empty:
            ral_de=pd.DataFrame({"Nro Etiqueta":ral_de["Nro Serie"], "Venta": "", "Categoria": "R&L D/E"})
            ral_de.to_excel("R&L DyE.xlsx",index=False)
            archivos_creados+=1
    #FALDA
        df_falda=df[df["Producto Desc"].str.startswith("FALDA")]
        df_falda=df_falda[["Nro Serie"]]
        if not df_falda.empty:
            df_falda_final=pd.DataFrame({"Nro Etiqueta":df_falda["Nro Serie"],"Venta":"","Categoria":"FALDA D/E"})
            df_falda_final.to_excel("Falda.xlsx",index=False)
            archivos_creados+=1
    #GRASA
        df_grasa=df[df["Producto Desc"]=="GRASA"]
        df_grasa=df_grasa[["Nro Serie"]]
        if not df_grasa.empty:
            df_grasa=pd.DataFrame({"Nro Etiqueta":df_grasa["Nro Serie"],"Venta":"","Categoria":"Fat"})
            df_grasa.to_excel("Grasa.xlsx",index=False)
            archivos_creados+=1
    #ASADO
        valores_asado=["ASADO CON HUESO CON MATAMBRE CON ENTRAÑA A 4 COSTILLAS (CAJA)","ASADO CON HUESO CON MATAMBRE CON ENTRAÑA A 5 COSTILLAS (CAJA)"]
        df_asado=df[df["Producto Desc"].isin(valores_asado)]
        df_asado=df_asado[["Nro Serie"]]
        if not df_asado.empty:
            df_asado=pd.DataFrame({"Nro Etiqueta":df_asado["Nro Serie"],"Venta":"","Categoria":"Asado D/E"})
            df_asado.to_excel("Asado D-E.xlsx",index=False)
            archivos_creados+=1
    #FFQCUPO y FFQDE
        valores_ffq=["AGUJA (CAJA)","CHINGOLO (CAJA)","COGOTE (CAJA)","MARUCHA (CAJA)","PECHO (CAJA)","PALETA (CAJA)"]
        df_filtrado_ffq=pd.DataFrame()
        for v in valores_ffq:
            filtroFFQ=df["Producto Desc"]==v
            df_filtrado_ffq=pd.concat([df_filtrado_ffq,df[filtroFFQ]],ignore_index=True)
        ffq_cupo=df_filtrado_ffq[pd.isna(df_filtrado_ffq["Contramarca"])]
        if not ffq_cupo.empty:
            ffq_cupo_final=pd.DataFrame({"Nro Etiqueta":ffq_cupo["Nro Serie"],"Venta":"","Categoria":"FFQ"})
            ffq_cupo_final.to_excel("FFQ CUPO.xlsx",index=False)
            archivos_creados+=1
        ffq_de_filtro=["D/E","T"]
        ffq_de= df_filtrado_ffq[df_filtrado_ffq["Contramarca"].isin(ffq_de_filtro)]
        if not ffq_de.empty:
            ffq_de_final=pd.DataFrame({"Nro Etiqueta":ffq_de["Nro Serie"],"Venta":"","Categoria":"FFQ - D/E"})
            ffq_de_final.to_excel("FFQ D-E.xlsx",index=False)
            archivos_creados+=1
    #90VLCUPO  y 90VLDE
        valores_90vl=["CUARTO TRASERO INCOMPLETO (CAJA)","CUARTO DELANTERO INCOMPLETO (CAJA)"]
        df_filtrado_90=pd.DataFrame()
        for vl in valores_90vl:
            filtro90=df["Producto Desc"]==vl
            df_filtrado_90=pd.concat([df_filtrado_90,df[filtro90]],ignore_index=True)
        df_90vl_cupo = df_filtrado_90[(df_filtrado_90["Contramarca"] == "GL") | (df_filtrado_90["Contramarca"].isna())]
        if not df_90vl_cupo.empty:
            df_90vl_cupo_final = pd.DataFrame({"Nro Etiqueta": df_90vl_cupo["Nro Serie"], "Venta": "", "Categoria": "Inc. 90 VL"})
            df_90vl_cupo_final.to_excel("90CUPO.xlsx", index=False)
            archivos_creados+=1
        de_90vl_filtro=["D/E","T"]
        df_90vl_de=df_filtrado_90[df_filtrado_90["Contramarca"].isin(de_90vl_filtro)]
        if not df_90vl_de.empty:
            df_90vl_de_final=pd.DataFrame({"Nro Etiqueta":df_90vl_de["Nro Serie"],"Venta":"","Categoria":"Inc. 90 VL D/E"})
            df_90vl_de_final.to_excel("90DyE.xlsx",index=False)
            archivos_creados+=1
    #RUEDACUPO y RUEDADE
        valores_rueda=["BOLA DE LOMO (CAJA)","CUADRADA (CAJA)","NALGA DE ADENTRO C/T (CAJA)","PECETO (CAJA)"]
        df_rueda=pd.DataFrame()
        for rueda in valores_rueda:
            filtro_rueda=df["Producto Desc"]==rueda
            df_rueda=pd.concat([df_rueda,df[filtro_rueda]],ignore_index=True)
        rueda_cupo=df_rueda[pd.isna(df_rueda["Contramarca"])]
        if not rueda_cupo.empty:
            rueda_cupo_final=pd.DataFrame({"Nro Etiqueta":rueda_cupo["Nro Serie"],"Venta":"","Categoria":"Rueda"})
            rueda_cupo_final.to_excel("Rueda Cupo.xlsx",index=False)
            archivos_creados+=1
        rueda_de_filtro=["D/E","T"]
        rueda_de=df_rueda[df_rueda["Contramarca"].isin(rueda_de_filtro)]
        if not rueda_de.empty:
            rueda_de_final=pd.DataFrame({"Nro Etiqueta":rueda_de["Nro Serie"],"Venta":"","Categoria":"RUEDA - D/E"})
            rueda_de_final.to_excel("Rueda DyE.xlsx",index=False)
            archivos_creados+=1
    #GTBCUPO y GTBDE
        valores_gtb=["BRAZUELO (CAJA)","GARRON (CAJA)","TORTUGUITA (CAJA)"]
        df_gtb=pd.DataFrame()
        for gtb in valores_gtb:
            filtro_gtb=df["Producto Desc"]==gtb
            df_gtb=pd.concat([df_gtb,df[filtro_gtb]],ignore_index=True)
        gtb_cupo=df_gtb[pd.isna(df_gtb["Contramarca"])]
        if not gtb_cupo.empty:
            gtb_cupo_final=pd.DataFrame({"Nro Etiqueta":gtb_cupo["Nro Serie"],"Venta":"","Categoria":"SSHM"})
            gtb_cupo_final.to_excel("GTB CUPO.xlsx",index=False)
            archivos_creados+=1
        gtb_de_filtro=["D/E","T"]
        gtb_de=df_gtb[df_gtb["Contramarca"].isin(gtb_de_filtro)]
        if not gtb_de.empty:
            gtb_de_final=pd.DataFrame({"Nro Etiqueta":gtb_de["Nro Serie"],"Venta":"","Categoria":"SSHM D/E"})
            gtb_de_final.to_excel("GTB DyE.xlsx",index=False)
            archivos_creados+=1
        if archivos_creados>0:
            print(f"Operación completada con éxito. Se crearon {archivos_creados} archivos.")
        else:
            print("No se crearon archivos. Verifique que el archivo de entrada contenga los datos necesarios.")
    except Exception:
        print(f"Error: Verifique la estructura del archivo {name}.")
        sys.exit()
if __name__ == "__main__":
    categorias()

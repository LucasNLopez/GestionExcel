import pandas as pd
import requests

def categorias ():
    #Lectura del excel
    df=pd.read_excel("C:\\Users\\llopez\\OneDrive - Frigorífico Alberdi S.A\\Escritorio\\categorias\\categorias.xlsx")
    
    #Recorte 70
    df_70 = df[df["Producto Desc"].str.contains("70 vl", case=False)]
    df_70=df_70[["Nro Serie","Contramarca"]]
    df_70VL = pd.DataFrame({"Nro Etiqueta": df_70["Nro Serie"], "Venta": "", "Categoria": "Trimming 70vl"})
    df_70VL.to_excel("70vl.xlsx", index=False)

    #Huesos
    df_huesos=df[df["Producto Desc"].str.startswith("HUESO")]
    df_huesos=df_huesos[["Nro Serie"]]
    df_hueso_final=pd.DataFrame({"Nro Etiqueta":df_huesos["Nro Serie"],"Venta":"","Categoria":"Mix Huesos"})
    df_hueso_final.to_excel("huesos.xlsx",index=False)
    
    #Mix Cupo y Mix D/E : Lista de valores que deseas buscar
    valores_a_buscar = ["BIFE ANCHO S/T-+2 kg. (CAJA) ", "BIFE ANCHO S/T-+ 2,5 (CAJA)", "BIFE ANCHO S/T--2 kg. (CAJA)",
                        "BIFE ANGOSTO CON CORDON-+3,5 kg. (CAJA) ","BIFE ANGOSTO CON CORDON-2,5 A 3 (CAJA)",
                        "BIFE ANGOSTO CON CORDON-3 A 3,5 (CAJA)","BIFE ANGOSTO CON CORDON-3 A 3.5 kg. (CAJA) ",
                        "BIFE ANGOSTO CON CORDON--3 kg. (CAJA) ","BIFE ANGOSTO CON CORDON-3,5 A 4 (CAJA)","BIFE ANGOSTO CON CORDON-4 A 5 (CAJA)",
                        "CUADRIL (CAJA)","GRASA DE PECHO (CAJA)","LOMO S/C +5 (CAJA)","LOMO S/C 2 LB (CAJA)","LOMO S/C 2/3 LB (CAJA)",
                        "LOMO S/C 3/4 LB (CAJA)","LOMO S/C 4/5 LB (CAJA)"]

    df_filtrado = pd.DataFrame()  # Crea un DataFrame vacío para almacenar los resultados

    for valor in valores_a_buscar:
        filtro = df["Producto Desc"] == valor
        df_filtrado = pd.concat([df_filtrado, df[filtro]], ignore_index=True)
    # Ahora df_filtrado contiene las filas donde "Producto Desc" es igual a cualquiera de los valores en valores_a_buscar
    
    mix_cupo=df_filtrado[pd.isna(df_filtrado["Contramarca"])]
    mix_cupo_final=pd.DataFrame({"Nro Etiqueta":mix_cupo["Nro Serie"],"Venta":"","Categoria":"Mix Cuts"})
    mix_cupo_final.to_excel("Mix Cupo.xlsx",index=False)
    
    mix_de_filtro=["D/E","T"]
    mix_de = df_filtrado[df_filtrado["Contramarca"].isin(mix_de_filtro)]
    mix_de_final=pd.DataFrame({"Nro Etiqueta":mix_de["Nro Serie"],"Venta":"","Categoria":"Mix Cuts D/E"})
    mix_de_final.to_excel("Mix D-E.xlsx",index=False)

    #FALDA
    df_falda=df[df["Producto Desc"].str.startswith("FALDA")]
    df_falda=df_falda[["Nro Serie"]]
    df_falda_final=pd.DataFrame({"Nro Etiqueta":df_falda["Nro Serie"],"Venta":"","Categoria":"FALDA D/E"})
    df_falda_final.to_excel("Falda.xlsx",index=False)

    #GRASA
    df_grasa=df[df["Producto Desc"]=="GRASA"]
    df_grasa=df_grasa[["Nro Serie"]]
    df_grasa_final=pd.DataFrame({"Nro Etiqueta":df_grasa["Nro Serie"],"Venta":"","Categoria":"Fat"})
    df_grasa_final.to_excel("Grasa.xlsx",index=False)

    #ASADO
    valores_asado=["ASADO CON HUESO CON MATAMBRE CON ENTRAÑA A 4 COSTILLAS (CAJA)","ASADO CON HUESO CON MATAMBRE CON ENTRAÑA A 5 COSTILLAS (CAJA)"]
    df_asado=df[df["Producto Desc"].isin(valores_asado)]
    df_asado=df_asado[["Nro Serie"]]
    df_asado_final=pd.DataFrame({"Nro Etiqueta":df_asado["Nro Serie"],"Venta":"","Categoria":"Asado D/E"})
    df_asado_final.to_excel("Asado D-E.xlsx",index=False)

    #FFQCUPO y FFQDE
    valores_ffq=["AGUJA (CAJA)","CHINGOLO (CAJA)","COGOTE (CAJA)","MARUCHA (CAJA)","PECHO (CAJA)","PALETA (CAJA)"]
    df_filtrado_ffq=pd.DataFrame()
    for v in valores_ffq:
        filtroFFQ=df["Producto Desc"]==v
        df_filtrado_ffq=pd.concat([df_filtrado_ffq,df[filtroFFQ]],ignore_index=True)
    #Ahora df_filtrado_ffq contiene todos los valores donde Producto Desc es igual a todos los valores de valores_ffq

    ffq_cupo=df_filtrado_ffq[pd.isna(df_filtrado_ffq["Contramarca"])]
    ffq_cupo_final=pd.DataFrame({"Nro Etiqueta":ffq_cupo["Nro Serie"],"Venta":"","Categoria":"FFQ"})
    ffq_cupo_final.to_excel("FFQ CUPO.xlsx",index=False)

    ffq_de_filtro=["D/E","T"]
    ffq_de= df_filtrado_ffq[df_filtrado_ffq["Contramarca"].isin(ffq_de_filtro)]
    ffq_de_final=pd.DataFrame({"Nro Etiqueta":ffq_de["Nro Serie"],"Venta":"","Categoria":"FFQ - D/E"})
    ffq_de_final.to_excel("FFQ D-E.xlsx",index=False)

    #90VLCUPO  y 90VLDE
    valores_90vl=["CUARTO TRASERO INCOMPLETO (CAJA)","CUARTO DELANTERO INCOMPLETO (CAJA)"]
    df_filtrado_90=pd.DataFrame()
    for vl in valores_90vl:
        filtro90=df["Producto Desc"]==vl
        df_filtrado_90=pd.concat([df_filtrado_90,df[filtro90]],ignore_index=True)
    #Ahora df_filtrado_90 contiene todos los valores donde Producto Desc es igual a todos los valores de valores_90vl

    
    df_90vl_cupo = df_filtrado_90[(df_filtrado_90["Contramarca"] == "GL") | (df_filtrado_90["Contramarca"].isna())]
    df_90vl_cupo_final = pd.DataFrame({"Nro Etiqueta": df_90vl_cupo["Nro Serie"], "Venta": "", "Categoria": "Inc. 90 VL"})
    df_90vl_cupo_final.to_excel("90CUPO.xlsx", index=False)


    de_90vl_filtro=["D/E","T"]
    df_90vl_de=df_filtrado_90[df_filtrado_90["Contramarca"].isin(de_90vl_filtro)]
    df_90vl_de_final=pd.DataFrame({"Nro Etiqueta":df_90vl_de["Nro Serie"],"Venta":"","Categoria":"Inc. 90 VL D/E"})
    df_90vl_de_final.to_excel("90DyE.xlsx",index=False)

    #RUEDACUPO y RUEDADE
    valores_rueda=["BOLA DE LOMO (CAJA)","CUADRADA (CAJA)","NALGA DE ADENTRO C/T (CAJA)","PECETO (CAJA)"]
    df_rueda=pd.DataFrame()
    for rueda in valores_rueda:
        filtro_rueda=df["Producto Desc"]==rueda
        df_rueda=pd.concat([df_rueda,df[filtro_rueda]],ignore_index=True)
    #Ahora df_rueda contiene todos los valores donde Producto desc es igual a todos los valores de valores_rueda

    rueda_cupo=df_rueda[pd.isna(df_rueda["Contramarca"])]
    rueda_cupo_final=pd.DataFrame({"Nro Etiqueta":rueda_cupo["Nro Serie"],"Venta":"","Categoria":"Rueda"})
    rueda_cupo_final.to_excel("Rueda Cupo.xlsx",index=False)

    rueda_de_filtro=["D/E","T"]
    rueda_de=df_rueda[df_rueda["Contramarca"].isin(rueda_de_filtro)]
    rueda_de_final=pd.DataFrame({"Nro Etiqueta":rueda_de["Nro Serie"],"Venta":"","Categoria":"RUEDA - D/E"})
    rueda_de_final.to_excel("Rueda DyE.xlsx",index=False)

    #GTBCUPO y GTBDE
    valores_gtb=["BRAZUELO (CAJA)","GARRON (CAJA)","TORTUGUITA (CAJA)"]
    df_gtb=pd.DataFrame()
    for gtb in valores_gtb:
        filtro_gtb=df["Producto Desc"]==gtb
        df_gtb=pd.concat([df_gtb,df[filtro_gtb]],ignore_index=True)
    #Ahora df_gtb contiene todos los valores donde Producto Desc es igual a todos los valores de valores_gtb

    gtb_cupo=df_gtb[pd.isna(df_gtb["Contramarca"])]
    gtb_cupo_final=pd.DataFrame({"Nro Etiqueta":gtb_cupo["Nro Serie"],"Venta":"","Categoria":"SSHM"})
    gtb_cupo_final.to_excel("GTB CUPO.xlsx",index=False)

    gtb_de_filtro=["D/E","T"]
    gtb_de=df_gtb[df_gtb["Contramarca"].isin(gtb_de_filtro)]
    gtb_de_final=pd.DataFrame({"Nro Etiqueta":gtb_de["Nro Serie"],"Venta":"","Categoria":"SSHM D/E"})
    gtb_de_final.to_excel("GTB DyE.xlsx",index=False)
    



if __name__ == "__main__":
    categorias()
    
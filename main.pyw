import customtkinter as ctk
import tkinter as tk
from datetime import datetime
import tkinter.messagebox as mbox
from functions import *
from PIL import Image


raiz = ctk.CTk()

raiz._set_appearance_mode("light")
raiz.title("Portal Objetivos") #Titulo ventana
raiz.resizable(False,False) #Permita ajustar tamaño
raiz.geometry("650x350") #Ajuste de tamaño



ruta_archivo = "./Data/InfoRegion/Info_Reg_und.xlsx"
df = cargar_datos_excel(ruta_archivo)
regiones = obtener_regiones(df)
fecha_actual = datetime.now().strftime('%Y-%m-%d')
columnas_need = ["REGION","ZONA","CODIGO SUCURSAL","FECHA DE LA NOVEDAD*D/M/A","NOVEDAD","CODIGO VENDEDOR","GENERAL","ORDINARIO","PLUS","GOLD"]


def actualizar_zonas(region, df, zona_menu):
    zonas = obtener_zonas_por_region(df, region)
    zona_menu.configure(values=zonas)
    if zonas:
        zona_menu.set(zonas[0])
    else:
        zona_menu.set("Sin zonas")

def manejar_cargar_archivo():
    df = cargar_archivo()
    if df is not None:
        df.columns = [col.strip() for col in df.columns]
        print("Columnas detectadas:", df.columns.tolist())
        if set(columnas_need).issubset(df.columns):
            labelcargaarc.configure(text="Archivo válido. ¡Listo para subir!")
            global df_archivo_usuario
            df_archivo_usuario = df
            botonDrive.configure(state="normal")
        else:
            labelcargaarc.configure(text="El archivo no contiene las columnas requeridas.")
            botonDrive.configure(state="disable")
    else:
        labelcargaarc.configure(text="No se seleccionó ningún archivo.")

def subir_a_drive():
    try:
        agregar_filas_a_drive(df_archivo_usuario)
        # ✅ Guardar archivo temporal
        try:
            gauth = GoogleAuth()
            gauth.settings['service_config'] = {
                "client_json_file_path": resource_path("service_account.json"),
                "client_user_email": "undobjetivos@organic-acronym-457612-b3.iam.gserviceaccount.com"
            }
            try:
                nombre_archivo =str(MenuRegion.get())+"_"+str(zona_menu.get())+"_"+str(numero.get())+"_"+str(fecha_actual)+".xlsx"
                gauth.ServiceAuth()
                drive = GoogleDrive(gauth)
                folder_id = "1buB0IPVGi2p47UorMRgiJ0RCNdyJ4Nw3"
                ruta_local = guardar_excel_local(df_archivo_usuario,nombre_archivo)
                backup_id = guardar_copia_seguridad(resource_path("./Data/Temp/"+ ruta_local),nombre_archivo,folder_id,drive)
            except:
                print("Error al intentar subir copia de seguridad")
        except Exception as e:
            print("❌ Error al subir copia de seguridad:", str(e))
        mbox.showinfo("Éxito", "Datos agregados exitosamente.")
    except Exception as e:
        mbox.showerror("Error", str(e))
        
imagen = ctk.CTkImage(Image.open(resource_path("Multimedia/logo.png")), size=(250, 50))
label = ctk.CTkLabel(raiz, image=imagen, text="",corner_radius=50)  
label.pack(pady=25)

LabelRegion=ctk.CTkLabel(raiz,text="Región")
LabelRegion.place(x=130,y=95)
MenuRegion=ctk.CTkOptionMenu(
    raiz,
    values=regiones,
    command=lambda seleccion: actualizar_zonas(seleccion, df, zona_menu)
    )
MenuRegion.set("Selecciona una región")
MenuRegion.place(x=130,y=120)


LabelPdv=ctk.CTkLabel(raiz, text="ZonaSupervisión")
LabelPdv.place(x=330,y=95)
zona_menu=ctk.CTkOptionMenu(raiz,values=["Selecciona una región"])
zona_menu.set("Selecciona una zona de supervisión")
zona_menu.place(x=330,y=120)

labelnumero = ctk.CTkLabel(raiz,text="Número Telefonico")
labelnumero.place(x=130,y=170)
numero = ctk.CTkEntry(raiz)
numero.place(x=130,y=195)

labelcargaarc = ctk.CTkLabel(raiz,text="Adjunta tu excel de objetivos")
labelcargaarc.place(x=330,y=170)
boton=ctk.CTkButton(raiz,text="Adjuntar Archivo", command=manejar_cargar_archivo)
boton.place(x=330,y=195)

botonDrive=ctk.CTkButton(raiz,text="Subir Objetivo",fg_color="black",hover_color="green", command=subir_a_drive)
botonDrive.place(x=230,y=270)
botonDrive.configure(state="disabled")

raiz.mainloop()
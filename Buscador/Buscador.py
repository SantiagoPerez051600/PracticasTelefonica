import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, Button
from tkinter import PhotoImage

class LabelConTitulo(tk.Frame):
    def __init__(self, parent, titulo, contenido, x, y, largo, alto, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        
        ColorLetra = "black"
        ColorObjeto = "#a1c2f9"

        self.config(bg=ColorObjeto, bd=2, relief=tk.SOLID)
        
        self.titulo_label = tk.Label(self, text=titulo, font=("Helvetica", 10, "bold"), bg=ColorObjeto, fg=ColorLetra)
        self.contenido_label = tk.Label(self, text=contenido, font=("Helvetica", 8), bg=ColorObjeto, fg=ColorLetra)
        self.boton_copiar = tk.Button(self, text="Copiar Texto", command=self.copiar_texto, bg=ColorObjeto, fg=ColorLetra)

        self.titulo_label.pack()
        self.contenido_label.pack()
        self.boton_copiar.pack()
        
        # Centramos el Frame
        self.place(x=x, y=y, width=largo, height=alto)

    def actualizar_contenido(self, nuevo_contenido):
        self.contenido_label.config(text=str(nuevo_contenido))

    def copiar_texto(self):
        texto = self.contenido_label.cget("text")
        self.clipboard_clear()
        self.clipboard_append(texto)
        self.update()
        messagebox.showinfo("Copiado", "Texto copiado al portapapeles.")

def Mensaje_emergentes(titulo,mensaje):
    messagebox.showinfo(titulo, mensaje)

def buscar():
    valor_busqueda = entrada.get()
    archivo_excel = "C:/Buscador/CLIENTES.xlsx"

    if not valor_busqueda:
        Mensaje_emergentes("ERROR", "Por favor ingrese el valor a consultar")
    else:
        try:
            df = pd.read_excel(archivo_excel)
            valor_busqueda = "".join(valor_busqueda.split())  
            # Filtrar el DataFrame por el valor de NIT (o cualquier otra columna)
            filas_encontradas = df[df['NIT_DIG'].astype(str) == str(valor_busqueda)]
            if not filas_encontradas.empty:
                # Obtener el primer índice de las filas encontradas (podría haber más de uno)
                indice = filas_encontradas.index[0]
                llenar_campos(filas_encontradas, indice)
            else:
                Mensaje_emergentes("Error!!", f"No se encontraron filas con el valor: {valor_busqueda}")
        except Exception as e:
            Mensaje_emergentes("!ERROR!", f"Error al leer el archivo: {e}")
            Mensaje_emergentes("¡ERROR!", "Asegúrese de que el archivo esté nombrado correctamente como 'CLIENTES.xlsx' y se encuentre en la siguiente ruta: C:/Buscador/")
           


def llenar_campos(filas_encontradas, indice):
    label1.actualizar_contenido(str(filas_encontradas.at[indice, "SEGMENTO "]))
    label2.actualizar_contenido(str(filas_encontradas.at[indice, "SUBSEGMENTO LOCAL"]))
    label3.actualizar_contenido(str(filas_encontradas.at[indice, "Id Salesforce"]))
    label4.actualizar_contenido(str(filas_encontradas.at[indice, "Cluster Nuevo"]))
    label5.actualizar_contenido(str(filas_encontradas.at[indice, "NOMBRE CLIENTE"]))
    label6.actualizar_contenido(str(filas_encontradas.at[indice, "CLIENTE PRINCIPAL HOLDING"]))
    label7.actualizar_contenido(str(filas_encontradas.at[indice, "NIT-C"]))
    label8.actualizar_contenido(str(filas_encontradas.at[indice, "Clasificación"]))
    label9.actualizar_contenido(str(filas_encontradas.at[indice, "ASESOR"]))
    label10.actualizar_contenido(str(filas_encontradas.at[indice, "EMAIL"]))
    label11.actualizar_contenido(str(filas_encontradas.at[indice, "MOBILEPHONE"]))

def vaciar_campos():
    label1.actualizar_contenido("")
    label2.actualizar_contenido("")
    label3.actualizar_contenido("")
    label4.actualizar_contenido("")
    label5.actualizar_contenido("")
    label6.actualizar_contenido("")
    label7.actualizar_contenido("")
    label8.actualizar_contenido("")
    label9.actualizar_contenido("")
    label10.actualizar_contenido("")
    label11.actualizar_contenido("")

# def validar_numero(P):
#     if P == "" or P.isdigit():
#         return True
#     return False


ventana = tk.Tk()
ventana.title("BUSCADOR DE CLIENTES")
ventana.configure(bg="#1699e3")
ventana.resizable(False, False)
ventana.geometry("750x550+300+70")
imagen_fondo = PhotoImage(file="C:/Buscador/fondo.gif")  
label_fondo = tk.Label(ventana, image=imagen_fondo)
label_fondo.place(x=0, y=0, relwidth=1, relheight=1) 


# Caja de texto para ingresar el valor de búsqueda
#vcmd = (ventana.register(validar_numero), '%P')  # Registrar la función de validación
entrada = tk.Entry(ventana, width=25, font=("Helvetica", 12), validate="key")
entrada.pack(pady=22)

# Botón de búsqueda
boton_buscar = tk.Button(ventana, text="Buscar", command=buscar, font=("Helvetica", 12), height=1)
boton_buscar.place(x=600, y=10)
boton_buscar.pack()
# Botón de borrar
boton_borrar = Button(ventana, text="Borrar Todo", command=vaciar_campos, font=("Helvetica", 12), height=1)
boton_borrar.place(x=640, y=15, width=90, height=30)

# Crear los labels con el mismo estilo y colocarlos uno al lado del otro
label1 = LabelConTitulo(ventana, "SEGMENTO ", "(VACIO)", 20, 210, 170, 80)
label2 = LabelConTitulo(ventana, "SUBSEGMENTO LOCAL", "(VACIO)",200, 210, 170, 80)
label3 = LabelConTitulo(ventana, "Id Salesforce", "(VACIO)", 380, 210, 170, 80)
label4 = LabelConTitulo(ventana, "Cluster Nuevo", "(VACIO)", 560, 210, 170, 80)
label5 = LabelConTitulo(ventana, "NOMBRE CLIENTE", "(VACIO)", 75, 120, 600, 80)
label6 = LabelConTitulo(ventana, "CLIENTE PRINCIPAL HOLDING", "(VACIO)", 20, 300, 300, 80)
label7 = LabelConTitulo(ventana, "NIT-C", "(VACIO)", 330, 300, 170, 80)
label8 = LabelConTitulo(ventana, "Clasificación", "(VACIO)", 510, 300, 220, 80)
label9 = LabelConTitulo(ventana, "ASESOR", "(VACIO)", 20, 390, 220, 80)
label10 = LabelConTitulo(ventana, "EMAIL", "(VACIO)", 250, 390, 300, 80)
label11 = LabelConTitulo(ventana, "MOBILEPHONE", "(VACIO)", 560, 390, 170, 80)

ventana.mainloop()

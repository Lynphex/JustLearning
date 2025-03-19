import pyodbc
import tkinter as tk
from tkinter import messagebox
import threading
import time


database_path = 'K:\Peticion gametos/Peticion gametos.accdb'
password = 'gensanta'
table_name_pattern = 'Peticiones pendientes'

def check_pending_requests():
    try:
        conn = pyodbc.connect(
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=' + database_path + ';'
            r'PWD=' + password
        )
        cursor = conn.cursor()

        # Obtiene los nombres de todas las tablas
        tables = cursor.tables(tableType='TABLE').fetchall()

        # Busca la tabla que coincide con el patrón
        pending_table = None
        for table in tables:
            if table_name_pattern in table.table_name:
                pending_table = table.table_name
                break

        if pending_table:
            # Ejecuta una consulta SQL para verificar si hay peticiones pendientes
            cursor.execute(f"SELECT COUNT(*) FROM [{pending_table}]")
            row = cursor.fetchone()
            
            if row[0] > 0:
                messagebox.showinfo("Resultado", f"Hay {row[0]} peticiones pendientes")
            else:
                messagebox.showinfo("Resultado", f"No hay peticiones pendientes")
        else:
            messagebox.showinfo("Resultado", f"No se encontró ninguna tabla que coincida con el patrón '{table_name_pattern}'.")

        # Cierra la conexión
        conn.close()

    except pyodbc.Error as e:
        messagebox.showerror("Error", f"Error al conectar con la base de datos: {e}")

# Función que se ejecuta cada 5 minutos
def periodic_check():
    while not stop_event.is_set():
        check_pending_requests()
        # Espera 5 minutos
        time.sleep(300)

# Función para iniciar el chequeo periódico
def start_periodic_check():
    threading.Thread(target=periodic_check).start()

def on_close():
    stop_event.set()
    root.destroy()

# Configuración de la interfaz gráfica
root = tk.Tk()
root.title("Monitor de Peticiones Pendientes")

# Botón para cerrar y salir
close_button = tk.Button(root, text="Cerrar y salir", command=on_close)
close_button.pack(pady=20)

# Evento para detener el hilo
stop_event = threading.Event()

# Inicia el chequeo periódico
start_periodic_check()

# Ejecuta el bucle principal de la interfaz gráfica
root.mainloop()
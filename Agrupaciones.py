import pandas as pd
import openpyxl
from itertools import combinations
import tkinter as tk
from tkinter import filedialog, messagebox

# Función para seleccionar un archivo Excel
def select_excel_file():
    filepath = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    return filepath

# Leer el archivo Excel y validar su formato
def read_excel_file(filepath):
    try:
        data = pd.read_excel(filepath)
        return data
    except Exception as e:
        raise ValueError(f"Error al leer el archivo Excel: {e}")

# Validar columnas requeridas
def validate_columns(data, required_columns):
    missing_columns = [col for col in required_columns if col not in data.columns]
    if missing_columns:
        raise ValueError(f"Faltan las siguientes columnas en el archivo Excel: {', '.join(missing_columns)}")

# Agrupar combinaciones aceptables con manejo preciso de circuitos y excedentes
def group_combinations(data, colors):
    grouped_results = []
    costums = []  # Lista para almacenar excedentes
    target_sums = [500, 1000]  # Solo 500 y 1000
    acceptable_ranges = {
        500: (450, 500),
        1000: (900, 1000),
    }

    for color in colors:
        col_name = f"Length {color}"
        values = data[["# de circuitos", col_name]].dropna().values
        sorted_values = sorted(values, key=lambda x: x[1], reverse=True)

        while sorted_values:
            current_group = []
            current_sum = 0
            indices_to_remove = []

            for i, (circuit, length) in enumerate(sorted_values):
                if current_sum + length <= 1000:  # Límite máximo
                    current_group.append((circuit, length))
                    current_sum += length
                    indices_to_remove.append(i)

                    if current_sum in target_sums:  # Si la suma alcanza exactamente 500 o 1000
                        grouped_results.append({
                            "Color": color,
                            "Combinación": [item[0] for item in current_group],
                            "Total": current_sum,
                            "Rollo": current_sum
                        })
                        break

            # Eliminar los circuitos utilizados
            sorted_values = [item for i, item in enumerate(sorted_values) if i not in indices_to_remove]

            # Manejar excedentes o grupos incompletos
            if current_sum not in target_sums and current_sum > 0:
                # Intentar optimizar excedentes combinando circuitos
                remaining_values = sorted([v[1] for v in sorted_values])
                for add_value in remaining_values:
                    if current_sum + add_value <= 500:  # Optimizar al límite de 500
                        current_sum += add_value
                        current_group.append((None, add_value))
                        sorted_values = [item for item in sorted_values if item[1] != add_value]
                        if current_sum == 500:
                            break

                # Almacenar el resultado final
                if current_sum in target_sums:
                    grouped_results.append({
                        "Color": color,
                        "Combinación": [item[0] for item in current_group if item[0] is not None],
                        "Total": current_sum,
                        "Rollo": current_sum
                    })
                else:
                    costums.append({
                        "Color": color,
                        "Combinación": [item[0] for item in current_group if item[0] is not None],
                        "Excedente": current_sum
                    })

        # Intentar fusionar excedentes para optimizar
        i = 0
        while i < len(costums):
            j = i + 1
            while j < len(costums):
                if costums[i]["Color"] == costums[j]["Color"]:
                    combined_sum = costums[i]["Excedente"] + costums[j]["Excedente"]
                    if combined_sum <= 500:
                        costums[i]["Combinación"] += costums[j]["Combinación"]
                        costums[i]["Excedente"] = combined_sum
                        costums.pop(j)
                        continue
                j += 1
            i += 1

    return grouped_results, costums

# Exportar resultados a Excel
def export_to_excel(grouped_results, costums, cable_size):
    file_path = filedialog.asksaveasfilename(
        title="Guardar archivo Excel",
        defaultextension=".xlsx",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )
    if file_path:
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            df_groups = pd.DataFrame(grouped_results)
            df_groups.to_excel(writer, sheet_name="Grupos", index=False)

            df_costums = pd.DataFrame(costums)
            df_costums.to_excel(writer, sheet_name="Costums", index=False)

        messagebox.showinfo("Éxito", f"Resultados exportados a {file_path}")
    else:
        messagebox.showwarning("Advertencia", "No se seleccionó ninguna ubicación para guardar el archivo.")

# Función principal de la interfaz gráfica
def main_gui():
    def process_data():
        try:
            # Obtener el archivo Excel
            filepath = select_excel_file()
            if not filepath:
                messagebox.showwarning("Advertencia", "No se seleccionó ningún archivo.")
                return

            # Leer datos
            data = read_excel_file(filepath)

            # Preguntar size del cable
            cable_size = cable_size_entry.get().strip()
            if not cable_size:
                messagebox.showwarning("Advertencia", "Debes ingresar el size del cable.")
                return

            # Determinar colores
            color_option = color_var.get()
            if color_option == 1:
                colors = ["Verde"]
            elif color_option == 2:
                colors = ["Negro", "Rojo", "Azul"]
            elif color_option == 3:
                colors = ["Negro", "Rojo", "Azul", "Blanco"]
            else:
                messagebox.showerror("Error", "Selección de colores inválida.")
                return

            # Validar columnas requeridas
            required_columns = ["# de circuitos"] + [f"Length {color}" for color in colors]
            validate_columns(data, required_columns)

            # Agrupar combinaciones
            grouped_results, costums = group_combinations(data, colors)

            if grouped_results or costums:
                export_to_excel(grouped_results, costums, cable_size)
            else:
                messagebox.showwarning("Sin resultados", "No se encontraron combinaciones aceptables.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # Crear la ventana principal
    root = tk.Tk()
    root.title("Agrupador de Cables")

    # Widgets para la entrada de datos
    tk.Label(root, text="Size del Cable:").grid(row=0, column=0, pady=5, padx=5)
    cable_size_entry = tk.Entry(root)
    cable_size_entry.grid(row=0, column=1, pady=5, padx=5)

    tk.Label(root, text="Colores a usar:").grid(row=1, column=0, pady=5, padx=5)
    color_var = tk.IntVar()
    tk.Radiobutton(root, text="Verde (ground)", variable=color_var, value=1).grid(row=1, column=1, sticky="w")
    tk.Radiobutton(root, text="Negro, Rojo, Azul", variable=color_var, value=2).grid(row=2, column=1, sticky="w")
    tk.Radiobutton(root, text="Negro, Rojo, Azul, Blanco", variable=color_var, value=3).grid(row=3, column=1, sticky="w")

    # Botón para procesar datos
    tk.Button(root, text="Procesar Archivo", command=process_data).grid(row=4, columnspan=2, pady=10)

    # Iniciar la interfaz gráfica
    root.mainloop()

if __name__ == "__main__":
    main_gui()



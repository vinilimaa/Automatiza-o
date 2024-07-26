import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
import pythoncom
import os

# Função para calcular os limites do retângulo envolvente
def bounding(coordinates):
    x = coordinates[::2]
    y = coordinates[1::2]
    xmin, ymin, xmax, ymax = min(x), min(y), max(x), max(y)
    return [xmin, ymin, xmax, ymax]

# Função para criar um objeto de ponto para o AutoCAD
def create_point_object(x, y):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y))

# Função para realizar a plotagem no formato A4
def plot_to_pdf(folder_path, acad):
    doc = acad.ActiveDocument
    nom = doc.Name[:-4]
    inicio = 1

    try:
        doc.SelectionSets.Item("SS1").Delete()
    except:
        pass

    ssget1 = doc.SelectionSets.Add("SS1")
    ssget1.SelectOnScreen()

    total_selected = len(ssget1)
    file_format = "{:02d}-{:02d}"

    k = 0
    z = inicio - 1

    layout = doc.ActiveLayout
    layout.ConfigName = "DWG To PDF.pc3"
    layout.PaperUnits = 1
    layout.CenterPlot = True
    layout.UseStandardScale = False
    layout.StyleSheet = "CTB_ARA.ctb"
    layout.PlotType = 4
    layout.StandardScale = 0
    layout.PlotWithLineweights = False

    for entity in ssget1:
        name = entity.EntityName

        if name == 'AcDbPolyline':
            k += 1
            file_number = file_format.format(k + z, total_selected)
            plotfile = os.path.join(folder_path, f"{nom}-{file_number}.pdf")
            coordinates = entity.Coordinates
            bounds = bounding(coordinates)
            P1 = create_point_object(bounds[0], bounds[1])
            P2 = create_point_object(bounds[2], bounds[3])

            width = abs(bounds[2] - bounds[0])
            height = abs(bounds[3] - bounds[1])

            if width > height:
                layout.CanonicalMediaName = "ISO_full_bleed_A0_(841.00_x_1189.00_MM)"
            else:
                layout.CanonicalMediaName = "ISO_full_bleed_A1_(841.00_x_594.00_MM)"

            layout.SetWindowToPlot(P1, P2)
            layout.CenterPlot = True
            doc.Plot.QuietErrorMode = False
            layout.SetCustomScale(1, 1)
            doc.SetVariable('BACKGROUNDPLOT', 0)
            layout.PlotRotation = 1  # Change this to the desired plot rotation
            doc.Plot.PlotToFile(plotfile)

    messagebox.showinfo("Plotagem Concluída", "Todas as plotagens foram concluídas com sucesso.")

# Função para lidar com a seleção de pasta
def browse_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_var.set(folder_path)

# Função para iniciar a plotagem
def start_plotting():
    folder_path = folder_var.get()
    if folder_path:
        try:
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
            plot_to_pdf(folder_path, acad)
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante a plotagem: {str(e)}")
    else:
        messagebox.showerror("Erro", "Por favor, selecione uma pasta para salvar os arquivos PDF.")

# Configurações da janela principal
root = tk.Tk()
root.title("Manel Plot Tool")
root.geometry("300x250")
root.configure(bg='#1C1A27')

folder_var = tk.StringVar()

# Frame principal
frame = tk.Frame(root, bg='#1C1A27')
frame.pack(expand=True)

# Título
label_title = tk.Label(frame, text="Manel Plot Tool", font=('Noto Sans', 20, 'bold'), bg='#1C1A27', fg='#ECF0F1')
label_title.pack(pady=10)

# Label e entrada para selecionar pasta
label_folder = tk.Label(frame, text="Selecione a pasta:", font=('Noto Sans', 12), bg='#1C1A27', fg='#ECF0F1')
label_folder.pack(pady=5)

entry_folder = tk.Entry(frame, textvariable=folder_var, font=('Noto Sans', 8), width=30)
entry_folder.pack(pady=5, padx=10, ipady=3)

# Botão para navegar
button_browse = tk.Button(frame, text="Navegar", command=browse_folder, bg='#E8083E', fg='#ECF0F1', font=('Noto Sans', 12, 'bold'), width=15, height=2, activebackground='#02864A', activeforeground='#1C1A27', borderwidth=0, highlightthickness=0)
button_browse.pack(pady=10)

# Botão para iniciar a plotagem
button_start = tk.Button(frame, text="Iniciar Plotagem", command=start_plotting, bg='#E8083E', fg='#ECF0F1', font=('Noto Sans', 12, 'bold'), width=15, height=2, activebackground='#02864A', activeforeground='#1C1A27', borderwidth=0, highlightthickness=0)
button_start.pack()

root.mainloop()

import tkinter as tk
from tkinter import filedialog
import win32com.client
import pythoncom
import time

# Funções de apoio para bounding e SPOINT

def bounding(Lista):
    x = []
    y = []

    for i in range(0, len(Lista), 2):
        x.append(Lista[i])
        y.append(Lista[i + 1])

    xmin = min(x)
    ymin = min(y)
    xmax = max(x)
    ymax = max(y)
    borde = [xmin, ymin, xmax, ymax]

    return borde

def SPOINT(x, y):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y))

def plotear_a4(ruta):
    # Inicia o AutoCAD
    acad = win32com.client.Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    
    # Obtém o nome do documento e remove a extensão
    nombre = doc.Name
    nom = nombre[0:len(nombre) - 4]
    
    inicio = 1
   
    # Coloque a pasta onde deseja salvar os arquivos PDF
    try:
        doc.SelectionSets.Item("SS1").Delete()
    except:
        pass  # Trata a exceção se a seleção não existir ou não puder ser deletada
    
    ssget1 = doc.SelectionSets.Add("SS1")
    ssget1.SelectOnScreen()
    
    k = 0
    z = inicio - 1
    
    for entity in ssget1:  # Percorre todas as entidades selecionadas
        time.sleep(0.05)  # Aguarda para evitar erros de velocidade
        name = entity.EntityName
        
        if name == 'AcDbPolyline':  # Verifica se a entidade é uma polilinha
            k += 1
            plotfile = ruta + "\\" + nom + "-0" + str(k + z).strip() + ".pdf"
            enti_pl1 = entity
            m = entity.Coordinates
            n = bounding(m)
            P1 = SPOINT(n[0], n[1])
            P2 = SPOINT(n[2], n[3])

            # Configuração do layout para plotagem
            doc.ActiveLayout.ConfigName = "DWG To PDF.pc3"
            doc.ActiveLayout.CanonicalMediaName = "ISO_full_bleed_A0_(841.00_x_1189.00_MM)"
            doc.ActiveLayout.SetWindowToPlot(P1, P2)
            doc.ActiveLayout.PaperUnits = 1
            doc.ActiveLayout.CenterPlot = True
            doc.Plot.QuietErrorMode = False
            doc.ActiveLayout.UseStandardScale = False
            doc.ActiveLayout.SetCustomScale(1, 1)
            doc.SetVariable('BACKGROUNDPLOT', 0)
            doc.Regen(1)
            doc.ActiveLayout.CenterPlot = True
            doc.ActiveLayout.PlotRotation = 1
            doc.ActiveLayout.StyleSheet = "CTB_ARA.ctb"
            doc.ActiveLayout.PlotType = 4
            doc.ActiveLayout.StandardScale = 0
            doc.Plot.PlotToFile(plotfile)
            time.sleep(5)

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_var.set(folder_path)

def start_plotting():
    folder_path = folder_var.get()
    if folder_path:
        plotear_a4(folder_path)
        status_var.set("Processo de plotagem concluído.")
    else:
        status_var.set("Por favor, selecione uma pasta.")

# Cria a janela principal
root = tk.Tk()
root.title("Manel Plot Tool")

# Variável para armazenar o caminho da pasta
folder_var = tk.StringVar()

# Variável para exibir o status da plotagem
status_var = tk.StringVar()
status_var.set("Aguardando início da plotagem...")

# Criação dos widgets
label_folder = tk.Label(root, text="Selecione a pasta de arquivos:")
entry_folder = tk.Entry(root, textvariable=folder_var, width=50)
button_browse = tk.Button(root, text="Navegar", command=browse_folder)
button_start = tk.Button(root, text="Iniciar Plotagem", command=start_plotting)
label_status = tk.Label(root, textvariable=status_var)

# Posicionamento dos widgets na janela
label_folder.grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_folder.grid(row=0, column=1, padx=5, pady=5)
button_browse.grid(row=0, column=2, padx=5, pady=5)
button_start.grid(row=1, column=0, columnspan=3, padx=5, pady=5)
label_status.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

# Criação dos widgets
label_folder = tk.Label(root, text="Selecione a pasta de arquivos:")
entry_folder = tk.Entry(root, textvariable=folder_var, width=50)
button_browse = tk.Button(root, text="Navegar", command=browse_folder)
button_start = tk.Button(root, text="Iniciar Plotagem", command=start_plotting)
label_status = tk.Label(root, textvariable=status_var)
label_about = tk.Label(root, text="Powered by: Emanoel Vinícius", font=("Arial", 10, "italic"), fg="gray")

# Posicionamento dos widgets na janela
label_folder.grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_folder.grid(row=0, column=1, padx=5, pady=5)
button_browse.grid(row=0, column=2, padx=5, pady=5)
button_start.grid(row=1, column=0, columnspan=3, padx=5, pady=5)
label_status.grid(row=2, column=0, columnspan=3, padx=5, pady=5)
label_about.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky="e")


# Loop principal da aplicação
root.mainloop()

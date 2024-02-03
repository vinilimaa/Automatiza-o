## Desenvolvido por Emanoel Vinícius 
## Emanoelufpe@gmail.com


import tkinter as tk
from tkinter import filedialog
import win32com.client
import pythoncom
import os

def bounding(Lista):
    x = Lista[::2]
    y = Lista[1::2]

    xmin = min(x)
    ymin = min(y)
    xmax = max(x)
    ymax = max(y)
    borde = [xmin, ymin, xmax, ymax]

    return borde

def SPOINT(x, y):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y))

def plotear_a4(ruta):
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except:
        print("Erro: AutoCAD não está ativo.")
        return

    doc = acad.ActiveDocument
    
    nombre = doc.Name
    nom = nombre[:-4]
    
    inicio = 1
   
    try:
        doc.SelectionSets.Item("SS1").Delete()
    except:
        pass
    
    ssget1 = doc.SelectionSets.Add("SS1")
    ssget1.SelectOnScreen()
    
    total_selecionado = len(ssget1)
    formato_numero = "{:02d}-{:02d}"

    k = 0
    z = inicio - 1
    
    for i, entity in enumerate(ssget1):
        name = entity.EntityName
        
        if name == 'AcDbPolyline':
            k += 1
            numero_arquivo = formato_numero.format(k + z, total_selecionado)
            plotfile = os.path.join(ruta, f"{nom}-{numero_arquivo}.pdf")
            enti_pl1 = entity
            m = entity.Coordinates
            n = bounding(m)
            P1 = SPOINT(n[0], n[1])
            P2 = SPOINT(n[2], n[3])

            largura = abs(n[2] - n[0])
            altura = abs(n[3] - n[1])

            doc.ActiveLayout.ConfigName = "DWG To PDF.pc3"
            if largura > altura:
                doc.ActiveLayout.CanonicalMediaName = "ISO_full_bleed_A0_(841.00_x_1189.00_MM)"
            else:
                doc.ActiveLayout.CanonicalMediaName = "ISO_full_bleed_A1_(841.00_x_594.00_MM)"
            
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
            doc.ActiveLayout.PlotWithLineweights = False
            doc.Plot.PlotToFile(plotfile)

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

root = tk.Tk()
root.title("Manel Plot Tool")
root.configure(bg='#2C2F33')  # Cor de fundo cinza escuro


# Centralizar a janela
window_width = 700
window_height = 150
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_cordinate = int((screen_width/2) - (window_width/2))
y_cordinate = int((screen_height/2) - (window_height/2))
root.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))

folder_var = tk.StringVar()
status_var = tk.StringVar()
status_var.set("Aguardando início da plotagem...")

# Estilos para os widgets
button_style = {
    'bg': '#3700B3',  # Cor vermelha para o botão
    'fg': '#FFFFFF',  # Texto em branco
    'relief': 'flat'  # Botão sem relevo
}

label_style = {
    'font': ('Noto Sans', 12),
    'bg': '#2C2F33',  # Cor de fundo cinza escuro
    'fg': '#FFFFFF'   # Texto em branco
}

entry_style = {
    'font': ('Noto Sans', 10),
    'bg': '#FFFFFF',  # Cor de fundo branca para o Entry
    'fg': '#000000'   # Texto em preto
}

# Criar widgets
label_folder = tk.Label(root, text="Selecione a pasta de arquivos:", **label_style)
entry_folder = tk.Entry(root, textvariable=folder_var, width=50, **entry_style)
button_browse = tk.Button(root, text="Navegar", command=browse_folder, **button_style)
button_start = tk.Button(root, text="Iniciar Plotagem", command=start_plotting, **button_style)
label_status = tk.Label(root, textvariable=status_var, **label_style)

# Posicionamento dos widgets
label_folder.grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_folder.grid(row=0, column=1, padx=5, pady=5)
button_browse.grid(row=0, column=2, padx=5, pady=5)
button_start.grid(row=1, column=0, columnspan=3, padx=5, pady=5)
label_status.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()

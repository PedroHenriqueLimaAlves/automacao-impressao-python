import win32print
import os
import re

ARQUIVO_ENTRADA = r"C:\SeparacaoSIAC\entrada.txt"
IMPRESSORA = r"\\estoque\Estoque FX-890"

# Comandos de Velocidade e Estilo
ESC = b'\x1b'
RESET = ESC + b'@'
ULTRA_DRAFT = ESC + b'x\x01' + ESC + b'k\x00'
NEGRITO_ON = ESC + b'G'
NEGRITO_OFF = ESC + b'H'
EXPANDIDO_ON = ESC + b'W1'
EXPANDIDO_OFF = ESC + b'W0'

def imprimir(dados, impressora):
    try:
        printer = win32print.OpenPrinter(impressora)
        win32print.StartDocPrinter(printer, 1, ("Separacao_Geral", None, "RAW"))
        win32print.StartPagePrinter(printer)
        win32print.WritePrinter(printer, RESET + ULTRA_DRAFT + dados)
        win32print.EndPagePrinter(printer)
        win32print.EndDocPrinter(printer)
        win32print.ClosePrinter(printer)
    except Exception as e:
        print(f"Erro: {e}")

def processar():
    if not os.path.exists(ARQUIVO_ENTRADA):
        return

    cabos = []
    cabecalho = []
    numero_pedido = "N/A"
    
    with open(ARQUIVO_ENTRADA, "r", encoding="latin-1") as f:
        linhas = f.readlines()

    for linha in linhas:
        if "Pedido...:" in linha:
            try: numero_pedido = linha.split("Pedido...:")[1].split()[0]
            except: pass

        item_match = re.search(r"^\d{4}\s+\(\s+\)", linha)
        
        if item_match:
          
            
            partes = linha.upper().split()
            
            if "CABO" in partes:
                idx_cabo = partes.index("CABO")
                # O "CABO" só conta se for o in

import pandas as pd
import os
import pywhatkit as pwk
import time
import sys
import re

# Diretório do script
if getattr(sys, 'frozen', False):
    # Executável
    current_dir = os.path.dirname(sys.executable)
else:
    # Ambiente normal Python
    current_dir = os.path.dirname(os.path.abspath(__file__))

# Arquivo XLSX
arquivo = os.path.join(current_dir, 'pendencia.csv')

# Colunas
importar_colunas= ["Cliente (Organização)", "Produto solicitado (suporte) **", "Quantidade","Cliente: Telefone (Organização)", "1 - Código Logística Reversa (log)"] 
novas_colunas = {"Cliente (Organização)": "cliente", "Produto solicitado (suporte) **":"produto", "Quantidade": "quantidade", "Cliente: Telefone (Organização)":"telefone", "1 - Código Logística Reversa (log)":"codigo_logistica"}

# Importação do CSV; renomear colunas
df = pd.read_csv(arquivo, sep=',', encoding='utf-8', usecols=importar_colunas)
df = df.rename(columns=novas_colunas)

def formatar_telefone(telefone):
    # Separar múltiplos números
    numeros = telefone.split(',')
    telefones_formatados = []

    for numero in numeros:
        numero = re.sub(r'\D', '', numero)  # Remove tudo que não for dígito
        if len(numero) == 10:  # Formato fixo com DDD
            telefone_formatado = f"({numero[:2]}) {numero[2:6]}-{numero[6:]}"
        elif len(numero) == 11:  # Formato celular com DDD
            telefone_formatado = f"({numero[:2]}) {numero[2:7]}-{numero[7:]}"
        else:
            telefone_formatado = None  # Número inválido

        if telefone_formatado:
            telefones_formatados.append(telefone_formatado)

    return ', '.join(telefones_formatados)

def enviar_mensagens(clientes, simular=False):
    for nome_cliente, dados_cliente in clientes.groupby('cliente'):
        nome_cliente_formatado = nome_cliente.split()[0].capitalize()  # Só o primeiro nome, capitalizado
        
        # Extrair telefones e preferir celular
        telefones = formatar_telefone(dados_cliente['telefone'].iloc[0]).split(',')
        numero_cliente = None
        for telefone in telefones:
            telefone = telefone.strip()
            if len(telefone) >= 5 and telefone[4] == ' ':  # Verifica se há um espaço após o DDD
                numero = telefone[5:]
                if numero.startswith('9') or numero.startswith('8'):  # Preferir números de celular
                    numero_cliente = telefone
                    break

        if not numero_cliente:
            print(f"[Atenção] Nenhum número de celular disponível para: {nome_cliente}. Ignorando cliente.")
            continue

        # Construir a mensagem
        mensagem = f"Bom dia! Tudo bem?\n"
        mensagem += "Identificamos que há equipamentos pendentes de devolução vinculados ao seu contrato. Para evitar bloqueios no envio de novos equipamentos e no atendimento do suporte técnico, solicitamos que regularize essa pendência com urgência.\n\n"
        mensagem += "Pedimos a gentileza de responder a esta mensagem confirmando a devolução ou informando eventuais dúvidas sobre o processo.\n"
        mensagem += "Equipamentos pendentes:\n"
        
        for _, linha in dados_cliente.iterrows():
            produto = linha['produto']
            quantidade = linha['quantidade']
            codigo = linha['codigo_logistica']
            
            if not pd.isna(codigo):
               mensagem += f"- {quantidade} x {produto.capitalize()} - {codigo}\n" 
            else:
               mensagem += f"- {quantidade} x {produto.capitalize()}\n" 
            

        mensagem += "Contamos com sua colaboração para regularizar a situação e continuar desfrutando dos nossos serviços sem interrupções.\n"
        mensagem += "\nSe precisar de algo, estamos à disposição! 😊🍺"

        # Enviar ou simular envio
        if simular:
            print(f"[Simulação] Mensagem para: {numero_cliente} - {nome_cliente} \n{mensagem}\n")
        else:
            # Enviar mensagem
            try:
                print(f"Enviando mensagem para {nome_cliente} - ({numero_cliente})...")
                pwk.sendwhatmsg_instantly(f"+55 {numero_cliente}", mensagem, wait_time=10, tab_close=True)
                time.sleep(8)  # Timer
            except Exception as e:
                print(f"Falha ao enviar mensagem para{nome_cliente} ({numero_cliente}): {e}")
            
    print("\nTodas as mensagens foram enviadas!\n")
            
if __name__ == '__main__':
    while True:
        user_input = input("Deseja rodar o programa em modo simulação? (s/n): ").strip().lower()
        if user_input == 's':
            simular = True
            break
        elif user_input == 'n':
            simular = False
            break
        else:
            print("Entrada inválida, por favor responda com 's' (sim) ou 'n' (não).")

    enviar_mensagens(df, simular)


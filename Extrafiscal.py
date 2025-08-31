import os
import re
import tempfile
import shutil
from io import BytesIO
import pandas as pd
import streamlit as st
import pdfplumber
from fpdf import FPDF
from datetime import datetime
from PIL import Image

# =================== CONFIGURAÇÃO ===================
st.set_page_config(page_title="CREA-RJ", layout="wide", page_icon="")

# =================== FUNÇÕES AUXILIARES ===================
def criar_temp_dir():
    """Cria diretório temporário"""
    return tempfile.mkdtemp()

def limpar_temp_dir(temp_dir):
    """Remove diretório temporário"""
    shutil.rmtree(temp_dir, ignore_errors=True)

def is_empty_info(text):
    """Verifica se o texto indica informação ausente"""
    if not text or str(text).strip() == '':
        return True
    return bool(re.search(r'^(SEM|NAO|NÃO|NAO INFORMADO|SEM INFORMAÇÃO)\s*[A-Z]*\s*$', str(text).strip(), re.IGNORECASE))

def clean_text(text):
    """Limpa texto removendo espaços extras e normalizando"""
    if not text:
        return ''
    text = str(text).replace('\n', ' ').strip()
    return ' '.join(text.split())

def formatar_agente_fiscalizacao(texto):
    """Formata o agente de fiscalização para manter apenas número e primeiro nome"""
    if not texto:
        return ''
    
    # Extrai o número e nome (padrão: "1010 - CELINA")
    match = re.match(r'(\d+\s*-\s*)([A-Za-zÀ-ÿ\s]+)', texto)
    if match:
        numero = match.group(1).strip()
        nome_completo = match.group(2).strip()
        primeiro_nome = nome_completo.split()[0].capitalize()
        return f"{numero} {primeiro_nome}"
    return texto

def get_nome_completo_agente(texto):
    """Obtém o nome completo del agente de fiscalização"""
    if not texto:
        return ''
    
    # Extrai o nome completo (padrão: "1010 - CELINA")
    match = re.match(r'\d+\s*-\s*([A-Za-zÀ-ÿ\s]+)', texto)
    if match:
        return match.group(1).strip()
    return texto

def formatar_responsavel(texto):
    """Formata o responsável para manter apenas la sigla inicial"""
    if not texto:
        return ''
    
    # Extrai la sigla (padrão: "")
    partes = [part.strip() for part in texto.split('-') if part.strip()]
    if partes:
        return partes[0]
    return texto

def formatar_data_relatorio(texto):
    """Extrai apenas la data del campo Data Relatório"""
    if not texto:
        return ''
    
    # Remove qualquer texto após la data (padrão: "22/05/2025    Fato Gerador:")
    match = re.search(r'(\d{2}/\d{2}/\d{4})', texto)
    if match:
        return match.group(1)
    return texto

def extrair_numero_protocolo(texto):
    """Extrai apenas o número do protocolo del campo Fato Gerador"""
    if not texto:
        return ''
    
    # Padrão: "" ou similar
    match = re.search(r'(?:PROCESSO|PROTOCOLO)[/\s]*(\d+)', texto, re.IGNORECASE)
    if match:
        return match.group(1)
    return ''

def extrair_numero_autuacao(texto):
    """Extrai o número de autuação do texto da seção 04"""
    if not texto:
        return ''
    
    # Padrão: "" (case insensitive)
    match = re.search(r'AUTUA[ÇC]AO\s+(\d+)', texto, re.IGNORECASE)
    if match:
        return match.group(1)
    return ''

def extrair_rf_principal(texto):
    """Extrai o RF Principal do texto"""
    if not texto:
        return ''
    
    # Padrão: "" (com ou sem espaços, case insensitive)
    match = re.search(r'RF Principal\s*:\s*(\d+)', texto, re.IGNORECASE)
    if match:
        return match.group(1)
    return ''

def extrair_secao(texto, titulo_secao):
    """Extrai o conteúdo de uma seção específica del PDF"""
    padrao = re.compile(
        r'{}\s*(.*?)(?=\s*\*\d+\s*-\s*|\Z)'.format(re.escape(titulo_secao)), 
        re.DOTALL | re.IGNORECASE
    )
    match = padrao.search(texto)
    if match:
        conteudo = match.group(1).strip()
        return None if is_empty_info(conteudo) else conteudo
    return None

def verificar_oficio(texto):
    """Verifica se contém registros de ofício no texto (retorna 1 se sim, 0 se não)"""
    if not texto or is_empty_info(texto):
        return 0
    
    # Verifica se contém a palavra "Ofício" (com ou sem acento) e variações
    padroes = [
        r'of[ií]cio',
        r'of\.',
        r'ofc',
        r'oficio',
        r'of[\s\-]?[0-9]'
    ]
    
    texto_str = str(texto).lower()
    for padrao in padroes:
        if re.search(padrao, texto_str, re.IGNORECASE):
            return 1
    return 0

def verificar_resposta_oficio(texto):
    """Verifica se contém 'Cópia ART' no texto (retorna 1 se sim, 0 se não)"""
    if not texto or is_empty_info(texto):
        return 0
    
    # Verifica se contém "Cópia ART" (case insensitive)
    texto_str = str(texto).lower()
    if re.search(r'c[óo]pia\s+art', texto_str, re.IGNORECASE):
        return 1
    return 0

def encontrar_pagina_secao_fotos(texto_completo, pdf):
    """Encontra la página onde está a seção 08 - Fotos"""
    for page_num, page in enumerate(pdf.pages, 1):
        texto_pagina = page.extract_text() or ""
        if "08 - Fotos" in texto_pagina:
            return page_num
    return None

def extrair_fotos_secao(pdf_path, temp_dir, filename):
    """Extrai apenas as fotos da seção 08 - Fotos, ignorando logotipos e assinaturas"""
    fotos_extraidas = []
    pdf_name = os.path.splitext(filename)[0]
    fotos_dir = os.path.join(temp_dir, "fotos", pdf_name)
    os.makedirs(fotos_dir, exist_ok=True)
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Extrai texto completo para encontrar a seção de fotos
            texto_completo = "\n".join(page.extract_text() or "" for page in pdf.pages)
            
            # Encontra la página onde está a seção de fotos
            pagina_fotos = encontrar_pagina_secao_fotos(texto_completo, pdf)
            
            if pagina_fotos is None:
                return fotos_extraidas
                
            # Processa apenas la página onde está a seção de fotos
            pagina = pdf.pages[pagina_fotos - 1]
            
            # Verifica se há imagens na página
            if hasattr(pagina, 'images') and pagina.images:
                # Filtra apenas imagens que estão provavelmente na seção de fotos
                # (exclui logotipos e assinaturas que geralmente estão no topo ou rodapé)
                altura_pagina = pagina.height
                
                for img_idx, img in enumerate(pagina.images):
                    try:
                        # Calcula a posição vertical da imagem (para excluir cabeçalho/rodapé)
                        y_pos = img['top']
                        
                        # Exclui imagens muito próximas del topo (logotipos) ou rodapé (assinaturas)
                        if y_pos < altura_pagina * 0.15 or y_pos > altura_pagina * 0.85:
                            continue
                            
                        # Exclui imagens muito pequenas (ícones, selos)
                        if img['width'] < 100 or img['height'] < 100:
                            continue
                            
                        # Extrai a imagem
                        if 'stream' in img:
                            img_data = img['stream'].get_data()
                            if img_data:
                                # Salva a imagem
                                img_name = f"foto_{img_idx+1}.png"
                                img_path = os.path.join(fotos_dir, img_name)
                                
                                with open(img_path, "wb") as f:
                                    f.write(img_data)
                                
                                fotos_extraidas.append({
                                    'nome': img_name,
                                    'caminho': img_path,
                                    'pagina': pagina_fotos
                                })
                    except Exception as e:
                        st.error(f"Erro ao extrair imagem {img_idx+1}: {e}")
    except Exception as e:
        st.error(f"Erro ao abrir PDF para extração de imagens: {e}")
    
    return fotos_extraidas

# =================== MÓDULO DE EXTRAÇÃO ===================
def extrair_todos_dados(texto, filename, pdf_path, temp_dir):
    """Extrai todos os dados del PDF de forma estruturada"""
    dados = {
        'RF': '',  # Alterado de 'Número' para 'RF'
        'RF Principal': '',  # Nova coluna para RF Principal
        'Situação': '',
        'Fiscal': '',  # Alterado de 'Agente de Fiscalização' para 'Fiscal'
        'Supervisão': '',  # Alterado de 'Responsável' para 'Supervisão'
        'Data': '',  # Alterado de 'Data Relatório' para 'Data'
        'Fato Gerador': '',
        'Protocolo': '',  # Nova coluna para o número do protocolo
        'Tipo Visita': '',
        
        # Seção 01 - Endereço Empreendimento
        'Endereço Empreendimento - Latitude': '',
        'Endereço Empreendimento - Longitude': '',
        'Endereço Empreendimento - Endereço': '',
        'Endereço Empreendimento - Descritivo': '',
        
        # Seção 02 - Identificação do Contratante
        'Identificação do Contratante': '',
        
        # Seção 03 - Atividade Desenvolvida
        'Atividade Desenvolvida': '',
        
        # Seção 04 - Identificação dos Contratados/Responsáveis
        'Identificação dos Contratados/Responsáveis': '',
        'Autuação': '',  # Nova coluna para número de autuação
        
        # Seção 05 - Documentos Solicitados/Expedidos
        'Documentos Solicitados/Expedidos': '',
        'Ofício': 0,  # Nova coluna para verificar ofícios (0 ou 1)
        
        # Seção 06 - Documentos Recebidos
        'Documentos Recebidos': '',
        'Resposta Ofício': 0,  # NOVA COLUNA: 1 se contém "Cópia ART", 0 se não
        
        # Seção 07 - Outras Informações
        'Outras Informações - Data Relatório Anterior': '',
        'Outras Informações - Informações Complementares': '',
        
        # Seção 08 - Fotos
        'Fotos': '',
        
        # Campos calculados
        'Ações': 0,
        
        # Campos adicionais para o relatório
        'Fiscal Nome Completo': '',  # Nome completo del agente de fiscalização
        'Supervisão Sigla': 'SBXD',   # Sigla fixa da supervisão
        
        # Informações para link de fotos
        'Nome Arquivo': filename,
        'Fotos Extraídas': 0
    }
    
    # Extrai metadados básicos
    campos_meta = [
        ('RF', r'Número\s*:\s*([^\n]+)'),  # Alterado de 'Número' para 'RF'
        ('Situação', r'Situação\s*:\s*([^\n]+)'),
        ('Fiscal', r'Agente\s+de\s+Fiscalização\s*:\s*([^\n]+)'),  # Alterado de 'Agente de Fiscalização' para 'Fiscal'
        ('Supervisão', r'Responsável\s*:\s*([^\n]+)'),  # Alterado de 'Responsável' para 'Supervisão'
        ('Data', r'Data\s+Relatório\s*:\s*([^\n]+)'),  # Alterado de 'Data Relatório' para 'Data'
        ('Fato Gerador', r'Fato\s+Gerador\s*:\s*([^\n]+)'),
        ('Protocolo', r'Protocolo\s*:\s*([^\n]+)'),
        ('Tipo Visita', r'Tipo\s+Visita\s*:\s*([^\n]+)')
    ]
    
    for campo, padrao in campos_meta:
        match = re.search(padrao, texto)
        if match:
            dados[campo] = clean_text(match.group(1))
    
    # Extrai o número do protocolo del campo Fato Gerador
    dados['Protocolo'] = extrair_numero_protocolo(dados['Fato Gerador'])
    
    # Extrai o RF Principal - CORREÇÃO: Busca em todo o texto, não apenas no campo específico
    dados['RF Principal'] = extrair_rf_principal(texto)
    
    # Aplica formatações específicas
    dados['Fiscal'] = formatar_agente_fiscalizacao(dados['Fiscal'])  # Alterado de 'Agente de Fiscalização' para 'Fiscal'
    dados['Supervisão'] = formatar_responsavel(dados['Supervisão'])  # Alterado de 'Responsável' para 'Supervisão'
    dados['Data'] = formatar_data_relatorio(dados['Data'])  # Alterado de 'Data Relatório' para 'Data'
    dados['Fiscal Nome Completo'] = get_nome_completo_agente(dados['Fiscal'])
    
    # Seção 01 - Endereço Empreendimento
    secao_endereco = extrair_secao(texto, "01 - Endereço Empreendimento")
    if secao_endereco:
        # Extrai latitude e longitude
        lat_match = re.search(r'Latitude\s*:\s*([-\d,.]+)', secao_endereco)
        long_match = re.search(r'Longitude\s*:\s*([-\d,.]+)', secao_endereco)
        if lat_match:
            dados['Endereço Empreendimento - Latitude'] = clean_text(lat_match.group(1))
        if long_match:
            dados['Endereço Empreendimento - Longitude'] = clean_text(long_match.group(1))
        
        # Extrai endereço (linha após coordenadas)
        coord_pattern = r'Latitude\s*:\s*[-\d,.]+\s*Longitude\s*:\s*[-\d,.]+'
        endereco_part = re.sub(coord_pattern, '', secao_endereco, flags=re.IGNORECASE)
        endereco_lines = [line.strip() for line in endereco_part.split('\n') if line.strip()]
        if endereco_lines:
            dados['Endereço Empreendimento - Endereço'] = clean_text(endereco_lines[0])
        
        # Extrai descritivo (se existir)
        if 'Descritivo:' in secao_endereco:
            desc_part = secao_endereco.split('Descritivo:')[-1]
            desc_text = clean_text(desc_part)
            if desc_text:
                dados['Endereço Empreendimento - Descritivo'] = desc_text
    
    # Seção 02 - Identificação do Contratante
    secao_contratante = extrair_secao(texto, "02 - Identificação del Contratante del Empreendimento")
    if secao_contratante:
        dados['Identificação do Contratante'] = clean_text(secao_contratante)
    
    # Seção 03 - Atividade Desenvolvida
    secao_atividade = extrair_secao(texto, "03 - Atividade Desenvolvida")
    if secao_atividade:
        dados['Atividade Desenvolvida'] = clean_text(secao_atividade)
    
    # Seção 04 - Identificação dos Contratados/Responsáveis
    secao_contratados = extrair_secao(texto, "04 - Identificação dos Contratados, Responsáveis Técnicos e/ou Fiscalizados")
    if secao_contratados:
        dados['Identificação dos Contratados/Responsáveis'] = clean_text(secao_contratados)
        # Extrai número de autuação se existir
        dados['Autuação'] = extrair_numero_autuacao(secao_contratados)
        
        # ALTERAÇÃO SOLICITADA: Calcula ações baseado em "Ramo Atividade" em vez de "Contratado" e "Responsável Técnico"
        ramos_atividade = len(re.findall(r'Ramo\s+Atividade\s*:', secao_contratados, re.IGNORECASE))
        dados['Ações'] = ramos_atividade
    
    # Seção 05 - Documentos Solicitados/Expedidos (MODIFICADA para pegar apenas conteúdo antes de "Fonte Informação")
    secao_docs_solicitados = extrair_secao(texto, "05 - Documentos Solicitados / Expedidos")
    if secao_docs_solicitados:
        # Divide o texto na primeira ocorrência de "Fonte Informação" e pega apenas a parte antes
        conteudo = secao_docs_solicitados.split("Fonte Informação")[0].strip()
        dados['Documentos Solicitados/Expedidos'] = clean_text(conteudo)
        # Verifica se contém ofícios (retorna 1 se sim, 0 se não)
        dados['Ofício'] = verificar_oficio(conteudo)
    
    # Seção 06 - Documentos Recebidos
    secao_docs_recebidos = extrair_secao(texto, "06 - Documentos Recebidos")
    if secao_docs_recebidos:
        dados['Documentos Recebidos'] = clean_text(secao_docs_recebidos)
        # Verifica se contém "Cópia ART" (retorna 1 se sim, 0 se não)
        dados['Resposta Ofício'] = verificar_resposta_oficio(secao_docs_recebidos)
    
    # Seção 07 - Outras Informações
    secao_outras = extrair_secao(texto, "07 - Outras Informações")
    if secao_outras:
        # Extrai data do relatório anterior
        data_anterior = re.search(r'Data\s+do\s+Relatório\s+Anterior\s*:\s*([^\n]+)', secao_outras)
        if data_anterior:
            dados['Outras Informações - Data Relatório Anterior'] = clean_text(data_anterior.group(1))
        
        # Extrai informações complementares
        info_complementares = re.search(r'Informações\s+Complementares\s*:\s*(.*)', secao_outras, re.DOTALL)
        if info_complementares:
            dados['Outras Informações - Informações Complementares'] = clean_text(info_complementares.group(1))
    
    # Seção 08 - Fotos (extrai as fotos do PDF)
    secao_fotos = extrair_secao(texto, "08 - Fotos")
    if secao_fotos:
        # Extrai as fotos apenas da seção 08 - Fotos
        fotos_extraidas = extrair_fotos_secao(pdf_path, temp_dir, filename)
        dados['Fotos Extraídas'] = len(fotos_extraidas)
        
        if fotos_extraidas:
            # Cria um link para as fotos no Excel
            dados['Fotos'] = f"{len(fotos_extraidas)} foto(s) extraída(s) da seção 08 - Fotos"
        else:
            dados['Fotos'] = "Seção de fotos encontrada, mas nenhuma imagem extraída"
    else:
        dados['Fotos'] = "Nenhuma seção de fotos encontrada"
    
    return dados

# =================== GERADORES DE RELATÓRIO ===================
def gerar_relatorio_completo(df):
    """Gera PDF com todos os dados extraídos com novo cabeçalho"""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Adiciona logo
    try:
        logo_path = "10.png"
        if os.path.exists(logo_path):
            pdf.image(logo_path, x=50, y=10, w=110)  # Centraliza a logo (ajuste as coordenadas conforme necessário)
            pdf.ln(40)  # Espaço após a logo
    except:
        pass
    
    # Título do relatório
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, 'Relatório Completo de Fiscalização', 0, 1, 'C')
    
    # Informações del agente e supervisão
    pdf.set_font('Arial', '', 12)
    
    # Obtém o nome completo del primeiro agente (assumindo que todos são do mesmo agente)
    nome_completo_agente = df.iloc[0]['Fiscal Nome Completo'] if 'Fiscal Nome Completo' in df.columns and len(df) > 0 else ''
    pdf.cell(0, 10, f'Agente de Fiscalização: {nome_completo_agente}', 0, 1)
    
    # Supervisão (fixo como SBXD)
    pdf.cell(0, 10, 'Supervisão: SBXD', 0, 1)
    
    # Período (primeira e última data)
    if len(df) > 0:
        datas = pd.to_datetime(df['Data'], errors='coerce', dayfirst=True)
        datas_validas = datas[~datas.isna()]
        if not datas_validas.empty:
            primeira_data = datas_validas.min().strftime('%d/%m/%Y')
            ultima_data = datas_validas.max().strftime('%d/%m/%Y')
            pdf.cell(0, 10, f'Período: {primeira_data} a {ultima_data}', 0, 1)
    
    # Data de geração do relatório
    pdf.cell(0, 10, f'Gerado em: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}', 0, 1)
    
    pdf.ln(10)
    
    # Configuração de colunas resumidas - ADICIONADAS COLUNAS RF PRINCIPAL, LEGALIZAÇÃO, FOTOS E RESPOSTA OFÍCIOS
    colunas = ['RFs', 'RF Principal', 'Legalização', 'Data', 'Ações', 'Ofícios', 'Resposta Ofícios', 'Protocolos', 'Autuações', 'Fotos']
    col_widths = [20, 25, 18, 15, 10, 10, 24, 15, 15, 10]
    
    # Cabeçalho
    pdf.set_font('Arial', 'B', 7)
    for w, h in zip(col_widths, colunas):
        pdf.cell(w, 8, h, 1, 0, 'C')
    pdf.ln()
    
    # Dados resumidos
    pdf.set_font('Arial', '', 7)
    
    # Conta o número de registros válidos (excluindo a linha de TOTAL se existir)
    df_validos = df[df['RF'] != 'TOTAL'] if 'TOTAL' in df['RF'].values else df
    num_registros = len(df_validos)
    
    # Variáveis para calcular totais
    total_acoes = 0
    total_oficios = 0
    total_resposta_oficios = 0
    total_protocolos = 0
    total_autuacoes = 0
    total_fotos = 0
    
    for _, row in df_validos.iterrows():
        # RF
        rf_text = str(row['RF'])[:15] + '...' if len(str(row['RF'])) > 15 else str(row['RF'])
        pdf.cell(col_widths[0], 8, rf_text, 1, 0, 'C')
        
        # RF Principal
        rf_principal_text = str(row['RF Principal'])[:15] + '...' if len(str(row['RF Principal'])) > 15 else str(row['RF Principal'])
        pdf.cell(col_widths[1], 8, rf_principal_text, 1, 0, 'C')
        
        # Legalização - 1 se tiver RF Principal, 0 se não tiver
        tem_legalizacao = '1' if row['RF Principal'] and str(row['RF Principal']).strip() != '' else '0'
        pdf.cell(col_widths[2], 8, tem_legalizacao, 1, 0, 'C')
        
        # Data
        pdf.cell(col_widths[3], 8, str(row['Data']), 1, 0, 'C')
        
        # Ações (agora baseado em Ramo Atividade)
        pdf.cell(col_widths[4], 8, str(row['Ações']), 1, 0, 'C')
        
        # Ofícios (usa a coluna Ofício que agora é 0 ou 1)
        pdf.cell(col_widths[5], 8, str(row['Ofício']), 1, 0, 'C')
        
        # Resposta Ofícios (usa a coluna Resposta Ofício que agora é 0 ou 1)
        pdf.cell(col_widths[6], 8, str(row['Resposta Ofício']), 1, 0, 'C')
        
        # Protocolos: 1 se tiver protocolo, 0 se não tiver
        tem_protocolo = '1' if row['Protocolo'] and str(row['Protocolo']).strip() != '' else '0'
        pdf.cell(col_widths[7], 8, tem_protocolo, 1, 0, 'C')
        
        # Autuações - 1 se tiver autuação, 0 se não tiver
        tem_autuacao = '1' if row['Autuação'] and str(row['Autuação']).strip() != '' else '0'
        pdf.cell(col_widths[8], 8, tem_autuacao, 1, 0, 'C')
        
        # Fotos - SIM se tem fotos extraídas, NÃO se não tem
        tem_fotos = 'SIM' if 'foto(s) extraída(s) da seção 08 - Fotos' in str(row['Fotos']) else 'NÃO'
        pdf.cell(col_widths[9], 8, tem_fotos, 1, 0, 'C')
        
        pdf.ln()
        
        # Acumula totais
        total_acoes += row['Ações'] if pd.notna(row['Ações']) else 0
        total_oficios += row['Ofício'] if pd.notna(row['Ofício']) else 0
        total_resposta_oficios += row['Resposta Ofício'] if pd.notna(row['Resposta Ofício']) else 0
        total_protocolos += 1 if tem_protocolo == '1' else 0
        total_autuacoes += 1 if tem_autuacao == '1' else 0
        total_fotos += 1 if tem_fotos == 'SIM' else 0
    
    # Linha de totais
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(col_widths[0], 8, f"TOTAL ({num_registros})", 1, 0, 'C')
    pdf.cell(col_widths[1], 8, "", 1, 0, 'C')  # RF Principal
    pdf.cell(col_widths[2], 8, "", 1, 0, 'C')  # Legalização
    pdf.cell(col_widths[3], 8, "", 1, 0, 'C')  # Data
    pdf.cell(col_widths[4], 8, str(total_acoes), 1, 0, 'C')  # Total Ações
    pdf.cell(col_widths[5], 8, str(total_oficios), 1, 0, 'C')  # Total Ofícios
    pdf.cell(col_widths[6], 8, str(total_resposta_oficios), 1, 0, 'C')  # Total Resposta Ofícios
    pdf.cell(col_widths[7], 8, str(total_protocolos), 1, 0, 'C')  # Total Protocolos
    pdf.cell(col_widths[8], 8, str(total_autuacoes), 1, 0, 'C')  # Total Autuações
    pdf.cell(col_widths[9], 8, str(total_fotos), 1, 0, 'C')  # Total Fotos
    pdf.ln()
    
    return pdf.output(dest='S').encode('latin1')

# =================== MÓDULO PRINCIPAL ===================
def extrator_pdf_consolidado():
    st.title("Lê os RFs, extrai os dados, gera planilha excel e produz Relatórios em PDF.")
    st.markdown("""
    **Extrai todos os dados dos PDFs para uma planilha Excel com formatação específica:**
    - Faz a leitura dos RFs em PDF e extrai todos os dados, gerando uma planilha excel.
    - Produz um relatório em PDF com os dados solicitados previamente.
    """)

    uploaded_files = st.file_uploader("Selecione os PDFs para extração", type="pdf", accept_multiple_files=True)
    
    if uploaded_files:
        temp_dir = criar_temp_dir()
        try:
            with st.spinner("Processando arquivos..."):
                dados_completos = []
                
                for file in uploaded_files:
                    temp_path = os.path.join(temp_dir, file.name)
                    with open(temp_path, "wb") as f:
                        f.write(file.getbuffer())
                    
                    # Extrai texto del PDF
                    with pdfplumber.open(temp_path) as pdf:
                        texto = "\n".join(page.extract_text() or "" for page in pdf.pages)
                    
                    # Processa com a função de extração completa (agora passando o caminho del PDF)
                    dados = extrair_todos_dados(texto, file.name, temp_path, temp_dir)
                    dados_completos.append(dados)
                    
                    os.unlink(temp_path)
                
                # Cria DataFrame com todos os dados
                df_completo = pd.DataFrame(dados_completos).fillna('')
                
                # Adiciona linha de totais
                df_total = pd.DataFrame([{
                    'RF': 'TOTAL',
                    'Ações': df_completo['Ações'].sum(),
                    'Ofício': df_completo['Ofício'].sum(),  # Soma dos ofícios (0s e 1s)
                    'Resposta Ofício': df_completo['Resposta Ofício'].sum()  # Soma das respostas de ofício (0s e 1s)
                }])
                df_completo = pd.concat([df_completo, df_total], ignore_index=True)
                
                # Reorganiza as colunas para colocar 'RF Principal' ao lado de 'Data'
                colunas = list(df_completo.columns)
                idx_data = colunas.index('Data')
                # Move 'RF Principal' para depois de 'Data'
                if 'RF Principal' in colunas:
                    colunas.insert(idx_data + 1, colunas.pop(colunas.index('RF Principal')))
                    df_completo = df_completo[colunas]
                
                # Exibe pré-visualização dos dados
                with st.expander("Visualizar dados extraídos", expanded=True):
                    st.dataframe(df_completo)
                
                # Gera relatório PDF
                pdf_completo = gerar_relatorio_completo(df_completo)
                
                # Download do relatório PDF
                st.success("Extração concluída com sucesso!")
                
                # Botão para baixar relatório PDF
                st.download_button(
                    "⬇️ Baixar Relatório de Fiscalização Completa",
                    pdf_completo,
                    "relatorio_completo.pdf"
                )
        
        finally:
            limpar_temp_dir(temp_dir)

# =================== INTERFACE PRINCIPAL ===================
def main():
    # Configuração visual
    try:
        logo = Image.open("10.png")
    except:
        logo = None
    
    # Layout do cabeçalho
    col1, col2 = st.columns([1, 2])
    with col1:
        if logo: st.image(logo, width=400)
    with col2:
        st.title("CREA-RJ - Conselho Regional de Engenharia e Agronomia do Rio de Janeiro")
    
    st.markdown("")
    
    # Exibe apenas o módulo principal
    extrator_pdf_consolidado()

    st.markdown("2025 - Carlos Franklin")

if __name__ == "__main__":
    main()
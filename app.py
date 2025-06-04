import os
import uuid # Para gerar nomes de arquivo únicos
import zipfile # Para criar arquivos ZIP
from flask import Flask, request, send_file, render_template_string, redirect, url_for, flash, g 
import pandas as pd
import pdfplumber
from collections import defaultdict

# --- Classe ExtratorTabelaPDF ---
class ExtratorTabelaPDF:
    """
    Classe responsável por extrair dados tabulares específicos de arquivos PDF.
    Foca nas colunas QTD, CÓDIGO e TITULO, lidando com títulos de múltiplas linhas.
    """
    # Constantes para nomes de colunas alvo
    COL_QTD = "QTD"
    COL_CODIGO = "CÓDIGO" # O script lida com a variação "CODIGO" internamente
    COL_TITULO = "TITULO"
    # Cabeçalhos usados como âncoras para definir limites de coluna
    COL_OPERACOES = "OPERAÇÕES"
    COL_FORNECIMENTO = "FORNECIMENTO"

    def __init__(self, debug=False, debug_image=False, base_path="."):
        """
        Inicializador da classe ExtratorTabelaPDF.

        :param debug: bool, se True, imprime mensagens de log detalhadas no console.
        :param debug_image: bool, se True, tenta salvar uma imagem da página PDF processada
                             com anotações de depuração (se a lógica de salvar imagem estiver habilitada).
        :param base_path: str, caminho base para salvar arquivos de depuração (como imagens).
        """
        self.debug = debug
        self.debug_image = debug_image
        self.base_path = base_path 

    def _print_debug(self, message):
        """Imprime mensagens de depuração se self.debug for True."""
        if self.debug:
            print(message)

    def _encontrar_limites_colunas_cabecalho(self, page_words, page_width_param, page_height_param):
        """
        Analisa as palavras extraídas de uma página PDF para encontrar os cabeçalhos
        das colunas de interesse e estimar seus limites horizontais (coordenadas x).

        Esta função é crucial para a extração baseada em texto, pois define as "faixas"
        onde o texto de cada coluna deve ser procurado.

        :param page_words: list, lista de dicionários, onde cada dicionário representa uma palavra
                           extraída da página com suas propriedades (texto, x0, x1, top, bottom).
        :param page_width_param: float, largura total da página PDF.
        :param page_height_param: float, altura total da página PDF.
        :return: tuple (found_headers_info, header_y_bottom_overall, column_boundaries)
                 found_headers_info: dict, informações sobre as palavras de cabeçalho encontradas.
                 header_y_bottom_overall: float, a coordenada Y inferior da linha de cabeçalho.
                 column_boundaries: dict, mapeia nomes de colunas para seus limites (x0, x1, y_top, y_bottom).
                 Retorna (None, 0, {}) se os cabeçalhos essenciais não forem encontrados.
        """
        found_headers_info = {}
        header_y_top_overall = page_height_param 
        header_y_bottom_overall = 0

        # Variantes de texto para cada cabeçalho que procuramos
        header_variants = {
            self.COL_OPERACOES: [self.COL_OPERACOES, "OPERACÕES", "OPERACOES"],
            self.COL_QTD: [self.COL_QTD, "QTD."],
            self.COL_CODIGO: [self.COL_CODIGO, "CODIGO", "CÓD.", "COD."],
            self.COL_TITULO: [self.COL_TITULO, "TÍTULO", "DESCRIÇÃO", "DESCRICAO", "TITULO"],
            self.COL_FORNECIMENTO: [self.COL_FORNECIMENTO] # Usado para delimitar o fim da coluna TITULO
        }
        
        possible_headers = defaultdict(list) # Armazena todas as palavras que correspondem a variantes de cabeçalho
        for word in page_words:
            current_word_text = word["text"].strip().upper() if word["text"] else ""
            for header_key, variants in header_variants.items():
                if current_word_text in [v.upper() for v in variants]:
                    possible_headers[header_key].append(word)
                    break # Palavra correspondeu a um tipo de cabeçalho, vai para a próxima palavra

        # Cabeçalhos essenciais para definir os limites das colunas de interesse
        essential_headers_for_boundaries = [self.COL_OPERACOES, self.COL_QTD, self.COL_CODIGO, self.COL_TITULO, self.COL_FORNECIMENTO]
        if not all(key in possible_headers for key in essential_headers_for_boundaries):
            self._print_debug(f"[DEBUG] Nem todos os cabeçalhos essenciais para limites ({essential_headers_for_boundaries}) foram encontrados.")
            self._print_debug(f"[DEBUG] Cabeçalhos encontrados por tipo: { {k: len(v) for k, v in possible_headers.items()} }")
            return None, 0, {}

        # Tenta encontrar a linha de cabeçalho mais coesa (palavras alinhadas verticalmente)
        # Usa o 'top' do primeiro cabeçalho essencial encontrado (na ordem definida) como referência vertical
        ref_y_top = 0
        for h_key in essential_headers_for_boundaries: 
            if possible_headers.get(h_key): # Se encontrou alguma palavra para este tipo de cabeçalho
                ref_y_top = min(w["top"] for w in possible_headers[h_key]) # Pega o 'top' mais alto entre elas
                self._print_debug(f"[DEBUG] Usando '{h_key}' (top: {ref_y_top}) como referência Y para a linha do cabeçalho.")
                break # Achou a referência, para
        if ref_y_top == 0: # Se nenhum dos cabeçalhos essenciais foi encontrado
            self._print_debug(f"[DEBUG] Nenhum cabeçalho essencial encontrado para ancorar a linha Y.")
            return None, 0, {}

        y_coord_tolerance = 5 # Palavras dentro desta tolerância vertical são consideradas na mesma linha de cabeçalho

        # Seleciona a melhor palavra (mais à esquerda e alinhada) para cada tipo de cabeçalho essencial
        for header_key in essential_headers_for_boundaries:
            best_word_for_header = None
            # Filtra as palavras candidatas para este header_key que estão alinhadas com ref_y_top
            aligned_candidates = [w for w in possible_headers.get(header_key, []) if abs(w["top"] - ref_y_top) < y_coord_tolerance]
            if aligned_candidates:
                best_word_for_header = min(aligned_candidates, key=lambda w: w["x0"]) # Pega a mais à esquerda
                found_headers_info[header_key] = {"word": best_word_for_header, 
                                                  "center_x": (best_word_for_header["x0"] + best_word_for_header["x1"]) / 2}
                header_y_top_overall = min(header_y_top_overall, best_word_for_header["top"])
                header_y_bottom_overall = max(header_y_bottom_overall, best_word_for_header["bottom"])
            else:
                # Se um cabeçalho essencial para definir os limites não for encontrado alinhado, é um problema.
                self._print_debug(f"[DEBUG] Cabeçalho essencial '{header_key}' não encontrado alinhado com ref_y_top={ref_y_top}. Falha na definição de limites.")
                return None, 0, {}

        # Define os limites das colunas (x0, x1) com base nas posições das palavras de cabeçalho encontradas
        column_boundaries = {}
        small_gap = 2  # Pequeno espaço entre o fim de uma coluna e o início da próxima
        col_min_width = 10 # Largura mínima para colunas estreitas como QTD ou CÓDIGO

        # Pega as informações das palavras de cabeçalho (já verificado que existem)
        op_info  = found_headers_info[self.COL_OPERACOES] 
        qtd_info = found_headers_info[self.COL_QTD]
        cod_info = found_headers_info[self.COL_CODIGO]
        tit_info = found_headers_info[self.COL_TITULO]
        forn_info = found_headers_info[self.COL_FORNECIMENTO] # FORNECIMENTO é usado para limitar TITULO

        # Coluna QTD: começa no x0 da palavra "QTD" e termina um pouco antes do x0 da palavra "CÓDIGO"
        qtd_x0 = qtd_info["word"]["x0"] - small_gap
        qtd_x1 = cod_info["word"]["x0"] - small_gap 
        column_boundaries[self.COL_QTD] = (max(0, qtd_x0), max(qtd_x0 + col_min_width, qtd_x1), header_y_top_overall, header_y_bottom_overall)

        # Coluna CÓDIGO: começa no x0 da palavra "CÓDIGO" e termina um pouco antes do x0 da palavra "TITULO"
        cod_x0 = cod_info["word"]["x0"] - small_gap
        cod_x1 = tit_info["word"]["x0"] - small_gap
        column_boundaries[self.COL_CODIGO] = (max(0, cod_x0), max(cod_x0 + col_min_width, cod_x1), header_y_top_overall, header_y_bottom_overall)
        
        # Coluna TITULO: começa no x0 da palavra "TITULO" e termina um pouco antes do x0 da palavra "FORNECIMENTO"
        tit_x0 = tit_info["word"]["x0"] - small_gap
        tit_x1 = forn_info["word"]["x0"] - small_gap # Delimitado pelo início de FORNECIMENTO
        column_boundaries[self.COL_TITULO] = (max(0, tit_x0), max(tit_x0 + 50, tit_x1), header_y_top_overall, header_y_bottom_overall) # Largura mínima de 50 para TITULO

        self._print_debug(f"[DEBUG] Cabeçalhos usados para limites: { {k:v['word']['text'] for k,v in found_headers_info.items()} }")
        self._print_debug(f"[DEBUG] Limites de coluna calculados (x0, x1, y_top_header, y_bottom_header): {column_boundaries}")
        self._print_debug(f"[DEBUG] Nível Y inferior do cabeçalho para iniciar a busca de dados: {header_y_bottom_overall}")
        
        return found_headers_info, header_y_bottom_overall, column_boundaries

    def _extrair_dados_baseado_em_texto(self, page):
        """
        Extrai os dados da tabela da página fornecida, usando uma abordagem baseada
        na análise de texto e coordenadas. Este é o método principal de extração.

        :param page: objeto Page do pdfplumber.
        :return: DataFrame do pandas com os dados extraídos (QTD, CÓDIGO, TITULO),
                 ou None se a extração falhar ou nenhum dado for encontrado.
        """
        # Extrai todas as palavras da página com tolerâncias justas para melhor agrupamento inicial
        words = page.extract_words(keep_blank_chars=False, use_text_flow=True, horizontal_ltr=True, 
                                   x_tolerance=1, y_tolerance=1) 
        if not words:
            self._print_debug("[DEBUG] Nenhuma palavra extraída da página.")
            return None

        page_width = page.width
        page_height = page.height
        self._print_debug(f"[DEBUG] Dimensões da página: Largura={page_width}, Altura={page_height}")

        # Encontra os cabeçalhos e define os limites das colunas de interesse
        headers_info, header_y_bottom_level, column_xbounds = \
            self._encontrar_limites_colunas_cabecalho(words, page_width, page_height)

        if not headers_info or not column_xbounds or not all(k in column_xbounds for k in [self.COL_QTD, self.COL_CODIGO, self.COL_TITULO]):
            self._print_debug("[DEBUG] Falha ao localizar cabeçalhos ou definir limites de coluna na extração principal.")
            return None

        # Filtra palavras que estão abaixo da linha do cabeçalho
        data_words = [w for w in words if w["top"] > header_y_bottom_level + 1] # +1 para um pequeno espaço
        
        # Define o limite Y inferior para parar de coletar dados (fim da tabela de itens)
        y_stop_limit = page_height 
        summary_line_anchors = ["Troca / R&I", "Troca/R&I", "TROCA / R&I", "TROCA/R&I"] # Variações da âncora
        
        candidate_summary_lines_y = []
        for word in data_words: # Procura a âncora nas palavras da área de dados
            cleaned_word_text = word["text"].strip() if word["text"] else ""
            if any(anchor.upper() in cleaned_word_text.upper() for anchor in summary_line_anchors):
                # A linha de resumo "Troca / R&I..." geralmente começa bem à esquerda
                if word["x0"] < page_width * 0.20: # Verifica se a palavra está na parte esquerda da página
                    candidate_summary_lines_y.append(word["top"])
        
        if candidate_summary_lines_y:
            y_stop_limit = min(candidate_summary_lines_y) - 2 # Pega a âncora mais alta e para um pouco antes
            self._print_debug(f"[DEBUG] Limite Y inferior para dados (antes do resumo da tabela de itens) definido em {y_stop_limit}.")
        else: 
            # Se a âncora principal não for encontrada, pode ser necessário um fallback,
            # mas para este caso, a ausência pode indicar que a tabela vai até o fim ou outro problema.
            self._print_debug(f"[DEBUG] Âncora de fim de tabela (ex: 'Troca / R&I') não encontrada na posição esperada. Processando até o fim da página ou próximo stop word.")
        
        data_words = [w for w in data_words if w["top"] < y_stop_limit] # Filtra palavras acima do limite de parada
        if not data_words:
            self._print_debug("[DEBUG] Nenhuma palavra de dados encontrada na área da tabela principal (abaixo do cabeçalho e antes do y_stop_limit).")
            return None

        # Agrupa palavras em "linhas candidatas" com base na proximidade vertical
        lines_raw = defaultdict(list)
        line_y_tolerance_raw = 2.5 # Tolerância vertical para agrupar palavras na mesma linha
        for word in data_words:
            word_v_center = (word["top"] + word["bottom"]) / 2 # Centro vertical da palavra
            matched_y_key = None
            for y_key in lines_raw.keys(): # y_key é o centro vertical da primeira palavra da linha
                if abs(word_v_center - y_key) < line_y_tolerance_raw:
                    matched_y_key = y_key
                    break
            if matched_y_key is None: matched_y_key = word_v_center # Cria uma nova linha
            lines_raw[matched_y_key].append(word)

        # Processa as linhas candidatas para montar as linhas da tabela e fundir títulos
        processed_rows = []
        sorted_y_keys_raw = sorted(lines_raw.keys()) # Processa linhas de cima para baixo
        
        for y_key in sorted_y_keys_raw:
            line_words = sorted(lines_raw[y_key], key=lambda w: w["x0"]) # Ordena palavras da linha por x0
            if not line_words: continue

            # Monta o texto para cada coluna nesta linha candidata
            current_row_assembly = {self.COL_QTD: [], self.COL_CODIGO: [], self.COL_TITULO: []}
            for word in line_words:
                word_text_clean = word["text"].strip() if word["text"] else ""
                if not word_text_clean: continue
                
                best_fit_col = None
                word_center_x = (word["x0"] + word["x1"]) / 2
                
                # Tenta encaixar a palavra na coluna baseada no centro ou sobreposição significativa
                for col_name, (x0_c, x1_c, _, _) in column_xbounds.items():
                    is_center_in_col = x0_c <= word_center_x < x1_c
                    
                    overlap_start = max(word["x0"], x0_c)
                    overlap_end = min(word["x1"], x1_c)
                    overlap_width = overlap_end - overlap_start
                    word_width = word["x1"] - word["x0"]
                    # Considera sobreposição significativa se for >40% da palavra ou >5 pixels absolutos
                    significant_overlap = (word_width > 0 and (overlap_width / word_width) > 0.4) or overlap_width > 5

                    if is_center_in_col or significant_overlap:
                        # Prioriza colunas mais à esquerda se houver ambiguidade de encaixe
                        if col_name == self.COL_QTD or col_name == self.COL_CODIGO:
                            if is_center_in_col or significant_overlap: # Condição mais forte
                               best_fit_col = col_name
                               break # Encaixou em QTD ou CODIGO, para
                        elif col_name == self.COL_TITULO: # TITULO é mais flexível
                            if best_fit_col is None : # Só atribui a TITULO se não encaixou melhor antes
                                best_fit_col = col_name 
                
                if best_fit_col:
                    current_row_assembly[best_fit_col].append(word_text_clean)
                # else:
                #     self._print_debug(f"[DEBUG] Palavra não atribuída: '{word_text_clean}' (x0:{word['x0']:.1f}, x1:{word['x1']:.1f}) Linha Y: {y_key:.1f}")


            # Cria um dicionário temporário com os dados da linha bruta
            temp_row_data = {
                self.COL_QTD: " ".join(current_row_assembly[self.COL_QTD]).strip(),
                self.COL_CODIGO: " ".join(current_row_assembly[self.COL_CODIGO]).strip(),
                self.COL_TITULO: " ".join(current_row_assembly[self.COL_TITULO]).strip()
            }
            # Adiciona à lista de processamento se tiver QTD ou TITULO (para permitir fusão posterior)
            if temp_row_data[self.COL_QTD] or temp_row_data[self.COL_TITULO]:
                processed_rows.append({"data": temp_row_data, "y_level": y_key})

        # Lógica para juntar títulos de múltiplas linhas
        merged_rows = []
        i = 0
        while i < len(processed_rows):
            current_item = processed_rows[i]["data"]
            current_y = processed_rows[i]["y_level"]
            
            is_current_qtd_valid = False
            actual_qtd_to_use = ""
            if current_item[self.COL_QTD]:
                try:
                    # Limpa QTD: pega o último token (geralmente o número "1") e tenta converter
                    parts_qtd = current_item[self.COL_QTD].replace(",", ".").strip().split(" ")
                    actual_qtd_to_use = parts_qtd[-1] 
                    float(actual_qtd_to_use) # Valida se é numérico
                    is_current_qtd_valid = True
                except ValueError:
                    is_current_qtd_valid = False

            # Se a linha atual tem um QTD válido, ela é o início de um item.
            # Tenta fundir com as próximas linhas se elas parecerem continuações do título.
            if is_current_qtd_valid:
                current_item[self.COL_QTD] = actual_qtd_to_use # Usa o QTD limpo
                j = i + 1
                while j < len(processed_rows):
                    next_item_data = processed_rows[j]["data"]
                    next_y = processed_rows[j]["y_level"]

                    # Verifica se a QTD da próxima linha está vazia ou não é numérica
                    is_next_qtd_empty_or_invalid = not next_item_data[self.COL_QTD] # Se QTD da próxima linha for vazio
                    if not is_next_qtd_empty_or_invalid: # Se QTD não for vazio, tenta validar
                        try:
                            parts_next_qtd = next_item_data[self.COL_QTD].replace(",", ".").strip().split(" ")
                            float(parts_next_qtd[-1])
                            is_next_qtd_empty_or_invalid = False # QTD é válido, então não é uma continuação do título
                        except ValueError:
                            is_next_qtd_empty_or_invalid = True # QTD não é numérico, PODE ser continuação

                    is_next_title_present = bool(next_item_data[self.COL_TITULO])
                    y_merge_tolerance = 15 # Tolerância vertical para considerar linhas como parte do mesmo item

                    # Condições para fusão: próxima linha não tem QTD válido, tem título, e está próxima verticalmente
                    if is_next_qtd_empty_or_invalid and is_next_title_present and (next_y - current_y) < y_merge_tolerance:
                        self._print_debug(f"[DEBUG] Juntando TITULO: '{current_item[self.COL_TITULO]}' + '{next_item_data[self.COL_TITULO]}'")
                        current_item[self.COL_TITULO] += " " + next_item_data[self.COL_TITULO]
                        current_item[self.COL_TITULO] = current_item[self.COL_TITULO].strip() # Remove espaços extras
                        current_y = next_y # Atualiza o y_level para a próxima comparação de fusão
                        j += 1
                    else:
                        break # Próxima linha não é uma continuação do título
                merged_rows.append(current_item)
                i = j # Pula para a próxima linha após as que foram fundidas
            else:
                # Se a linha atual não tem QTD válido, ela é descartada (a menos que seja fundida de alguma forma)
                self._print_debug(f"[DEBUG] Linha Inicial Descartada (QTD inválido/ausente '{current_item[self.COL_QTD]}' e não parte de fusão): {current_item}")
                i += 1
                
        if not merged_rows:
            self._print_debug("[DEBUG] Nenhuma linha de dados formatada após tentativa de fusão.")
            return None
            
        df_result = pd.DataFrame(merged_rows)
        if df_result.empty: return None

        # Garante a ordem correta das colunas e que todas as colunas de interesse estejam presentes
        final_columns_ordered = [self.COL_QTD, self.COL_CODIGO, self.COL_TITULO] 
        for col_name in final_columns_ordered:
            if col_name not in df_result.columns:
                df_result[col_name] = "" # Adiciona coluna vazia se não existir
        return df_result[final_columns_ordered]

    def _salvar_imagem_debug(self, page, pagina_num, caminho_base_pdf):
        """ Salva uma imagem de depuração da página PDF, se self.debug_image for True. """
        if not self.debug_image:
            return
        try:
            # Define o diretório para salvar a imagem de depuração
            dir_debug_img = self.base_path if self.base_path else os.path.dirname(caminho_base_pdf)
            if not dir_debug_img or dir_debug_img == ".": 
                 # Tenta obter o diretório absoluto do PDF se base_path não for útil
                 dir_debug_img = os.path.dirname(os.path.abspath(caminho_base_pdf))
                 if not os.path.exists(dir_debug_img) : dir_debug_img = "." # Fallback para diretório atual

            os.makedirs(dir_debug_img, exist_ok=True) # Garante que o diretório de depuração exista

            nome_base_pdf_img = os.path.splitext(os.path.basename(caminho_base_pdf))[0]
            path_img_debug = os.path.join(dir_debug_img, f"{nome_base_pdf_img}_pagina_{pagina_num}_visual_debug.png")
            
            page.to_image(resolution=150).save(path_img_debug, format="PNG")
            self._print_debug(f"[DEBUG] Imagem de depuração visual da página salva em: {path_img_debug}")
        except Exception as e_img_save:
            self._print_debug(f"[DEBUG] Não foi possível salvar imagem de depuração visual da página: {e_img_save}")

    def processar_pdf(self, caminho_pdf):
        """
        Processa o arquivo PDF fornecido, extrai a tabela de itens da primeira página.

        :param caminho_pdf: str, caminho para o arquivo PDF.
        :return: DataFrame do pandas com os dados extraídos, ou None se ocorrer um erro.
        """
        dataframe_resultado_final = None
        if not os.path.isfile(caminho_pdf):
            print(f"[ERRO] Arquivo PDF não encontrado: {caminho_pdf}")
            return None
        try:
            with pdfplumber.open(caminho_pdf) as pdf_doc:
                if pdf_doc.pages:
                    pagina_alvo = pdf_doc.pages[0] # Processa apenas a primeira página
                    self._print_debug(f"\n[DEBUG] Processando Página 1 de {len(pdf_doc.pages)} com extração baseada em texto.")
                    
                    # Salva imagem de depuração se a flag estiver ativa
                    self._salvar_imagem_debug(pagina_alvo, 1, caminho_pdf)
                    
                    dataframe_resultado_final = self._extrair_dados_baseado_em_texto(pagina_alvo)
                else:
                    self._print_debug("[ERRO] O PDF não contém páginas.")
            return dataframe_resultado_final
        except Exception as e_main:
            print(f"Erro CRÍTICO durante o processamento do PDF '{caminho_pdf}': {e_main}")
            import traceback
            traceback.print_exc() # Imprime o traceback completo para depuração
            return None

    def salvar_resultado_excel(self, dataframe, caminho_pdf_original, nome_arquivo_saida_opcional=None):
        """
        Salva o DataFrame fornecido em um arquivo Excel (.xlsx).

        :param dataframe: DataFrame do pandas a ser salvo.
        :param caminho_pdf_original: str, caminho do PDF original, usado para gerar o nome do arquivo de saída.
        :param nome_arquivo_saida_opcional: str, opcional, nome completo do arquivo de saída.
                                             Se None, um nome será gerado automaticamente.
        :return: str, caminho do arquivo Excel salvo, ou None se ocorrer um erro.
        """
        if dataframe is None or dataframe.empty:
            print("Nenhum dado para salvar em Excel.")
            return None 

        try:
            # Define o caminho de saída
            if nome_arquivo_saida_opcional:
                caminho_completo_saida = nome_arquivo_saida_opcional
                dir_saida = os.path.dirname(caminho_completo_saida)
                if not dir_saida: dir_saida = '.' # Diretório atual se apenas nome do arquivo for fornecido
            else:
                diretorio_pdf = os.path.dirname(caminho_pdf_original)
                if not diretorio_pdf: diretorio_pdf = '.' # Diretório atual se PDF está no mesmo dir do script
                
                nome_arquivo_base_pdf = os.path.splitext(os.path.basename(caminho_pdf_original))[0]
                subdiretorio_saida = "excel_saida" # Nome da pasta padrão para os Excels
                dir_saida = os.path.join(diretorio_pdf, subdiretorio_saida)
                caminho_completo_saida = os.path.join(dir_saida, f"{nome_arquivo_base_pdf}_extracao.xlsx")

            os.makedirs(dir_saida, exist_ok=True) # Cria o diretório de saída se não existir
            dataframe.to_excel(caminho_completo_saida, index=False, engine='openpyxl')
            print(f"Planilha Excel salva com sucesso em: {caminho_completo_saida}")
            return caminho_completo_saida 
        except Exception as e:
            print(f"Erro ao salvar o arquivo Excel: {e}")
            return None


# --- Configuração e Rotas do Flask App ---
app = Flask(__name__)
app.secret_key = os.urandom(24) # Chave secreta para mensagens flash e sessões

# Define pastas para upload e saída. Estas serão criadas no mesmo diretório do app.py.
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'arquivos_pdf_enviados') # Nome de pasta mais descritivo
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'arquivos_excel_gerados') # Nome de pasta mais descritivo
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
ALLOWED_EXTENSIONS = {'pdf'} # Apenas arquivos PDF são permitidos

def allowed_file(filename):
    """Verifica se a extensão do arquivo é permitida."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Template HTML para a página de upload (com Tailwind CSS - Cores Verdes)
HTML_FORM = """
<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
  <title>Extrator PDF para Excel</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <style>
    /* Estilos adicionais podem ser colocados aqui se necessário */
    body {
      font-family: 'Inter', sans-serif; /* Fonte padrão do Tailwind */
    }
  </style>
</head>
<body class="bg-gradient-to-br from-emerald-700 to-green-900 text-gray-100 flex items-center justify-center min-h-screen p-4">
  <div class="container bg-emerald-800 p-6 sm:p-8 md:p-10 rounded-xl shadow-2xl max-w-lg w-full">
    <header class="text-center mb-8">
      <!-- Ícone SVG de Upload -->
      <svg class="w-16 h-16 mx-auto mb-4 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path></svg>
      <h1 class="text-3xl sm:text-4xl font-bold text-transparent bg-clip-text bg-gradient-to-r from-emerald-400 to-green-300">Extrator de PDF para Excel</h1>
      <p class="text-emerald-300 mt-2">Faça upload de um ou mais arquivos PDF para extrair tabelas.</p>
    </header>
    
    <!-- Seção para exibir mensagens flash (erros, sucesso) -->
    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        <ul class="flash-messages mb-6 space-y-2">
        {% for category, message in messages %}
          <li class="{{ 'bg-red-600 border-red-800 text-red-100' if category == 'error' else 'bg-green-600 border-green-800 text-green-100' }} p-3 rounded-md shadow-sm text-sm">
            {{ message }}
          </li>
        {% endfor %}
        </ul>
      {% endif %}
    {% endwith %}
    
    <!-- Formulário de Upload -->
    <form method=post enctype=multipart/form-data class="space-y-6">
      <div>
        <label for="file-upload" class="block text-sm font-medium text-emerald-200 mb-1">Selecione os arquivos PDF:</label>
        <input 
          id="file-upload" 
          type=file 
          name=file  {# O nome do campo deve ser 'file' para request.files.getlist("file") #}
          accept=".pdf" 
          required 
          multiple {# Permite a seleção de múltiplos arquivos #}
          class="block w-full text-sm text-emerald-300
                 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0
                 file:text-sm file:font-semibold file:bg-emerald-600 file:text-emerald-50
                 hover:file:bg-emerald-700 cursor-pointer border border-emerald-600 rounded-lg p-2 focus:outline-none focus:ring-2 focus:ring-emerald-500"
        >
      </div>
      <input 
        type=submit 
        value="Processar Arquivos" 
        class="w-full px-6 py-3 bg-gradient-to-r from-emerald-500 to-green-500 text-white font-semibold rounded-lg shadow-md hover:from-emerald-600 hover:to-green-600 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:ring-opacity-75 cursor-pointer text-base transition-all duration-300 ease-in-out"
      >
    </form>
     <footer class="text-center mt-8 text-xs text-emerald-400">
        <p>&copy; 2024 Seu Aplicativo Extrator. Todos os direitos reservados.</p>
    </footer>
  </div>
</body>
</html>
"""

@app.after_request
def remove_temporary_files(response):
    """
    Função de limpeza executada após cada requisição.
    Remove os arquivos temporários (PDFs, XLSX, ZIP) que foram marcados para remoção.
    """
    files_to_remove_list = getattr(g, 'files_to_remove', [])
    for f_path in files_to_remove_list:
        try:
            if f_path and os.path.exists(f_path):
                os.remove(f_path)
                print(f"Arquivo temporário removido: {f_path}")
        except Exception as e:
            print(f"Erro ao remover arquivo temporário {f_path}: {e}")
    return response


@app.route('/', methods=['GET', 'POST'])
def rota_upload_arquivo(): 
    """
    Rota principal da aplicação.
    GET: Exibe o formulário de upload.
    POST: Processa os arquivos PDF enviados, extrai os dados, e oferece o(s) arquivo(s) Excel para download.
    """
    g.files_to_remove = [] # Inicializa a lista de arquivos para remoção para esta requisição
    
    if request.method == 'POST':
        # Pega a lista de arquivos enviados (o input HTML deve ter `multiple`)
        arquivos_enviados = request.files.getlist("file") 

        if not arquivos_enviados or not arquivos_enviados[0].filename: 
            flash('Nenhum arquivo selecionado. Por favor, escolha um ou mais PDFs.', 'error')
            return redirect(request.url)

        arquivos_excel_processados_info = [] # Lista para armazenar informações dos Excels gerados
        
        for arquivo_storage in arquivos_enviados: # Itera sobre cada arquivo enviado
            if arquivo_storage and allowed_file(arquivo_storage.filename):
                nome_arquivo_original = arquivo_storage.filename
                id_unico_arquivo = str(uuid.uuid4()) # ID único para este arquivo
                
                # Define caminhos para o PDF salvo e o Excel de saída
                nome_pdf_temporario = f"{id_unico_arquivo}.pdf"
                caminho_pdf_salvo = os.path.join(app.config['UPLOAD_FOLDER'], nome_pdf_temporario)
                g.files_to_remove.append(caminho_pdf_salvo) # Marca PDF para remoção
                
                try:
                    arquivo_storage.save(caminho_pdf_salvo) # Salva o PDF enviado
                    flash(f'Arquivo "{nome_arquivo_original}" recebido. Processando...', 'success')

                    # Instancia e usa o extrator
                    # debug=True para ver logs no console do Flask, debug_image=False para não salvar imagens no servidor
                    extrator = ExtratorTabelaPDF(debug=True, debug_image=False, base_path=app.config['UPLOAD_FOLDER']) 
                    df_extraido = extrator.processar_pdf(caminho_pdf_salvo)

                    if df_extraido is not None and not df_extraido.empty:
                        nome_excel_temporario = f"{id_unico_arquivo}_extracao.xlsx"
                        caminho_excel_saida_temp = os.path.join(app.config['OUTPUT_FOLDER'], nome_excel_temporario)
                        
                        caminho_excel_gerado = extrator.salvar_resultado_excel(df_extraido, 
                                                                              caminho_pdf_salvo, 
                                                                              nome_arquivo_saida_opcional=caminho_excel_saida_temp)
                        if caminho_excel_gerado and os.path.exists(caminho_excel_gerado):
                            arquivos_excel_processados_info.append({
                                "original_name": nome_arquivo_original, # Nome original do PDF
                                "excel_path": caminho_excel_gerado     # Caminho para o XLSX gerado
                            })
                            g.files_to_remove.append(caminho_excel_gerado) # Marca XLSX para remoção
                        else:
                            flash(f'Falha ao gerar o arquivo Excel para "{nome_arquivo_original}".', 'error')
                    else:
                        flash(f'Não foram extraídos dados do PDF "{nome_arquivo_original}" ou o resultado estava vazio.', 'error')
                except Exception as e_proc:
                    print(f"Erro ao processar o arquivo {nome_arquivo_original}: {e_proc}")
                    flash(f'Erro ao processar o arquivo "{nome_arquivo_original}". Verifique os logs do servidor.', 'error')
            elif arquivo_storage.filename: # Se tem nome mas não é um PDF permitido
                flash(f'Tipo de arquivo não permitido para "{arquivo_storage.filename}". Apenas PDFs são aceitos.', 'error')
        
        # Após processar todos os arquivos enviados
        if not arquivos_excel_processados_info: # Se nenhum arquivo foi processado com sucesso
            if not get_flashed_messages(category_filter=['error']): # Evita duplicar msg se já houve erro específico
                 flash('Nenhum arquivo PDF foi processado com sucesso ou nenhum dado foi extraído.', 'error')
            return redirect(request.url)

        if len(arquivos_excel_processados_info) == 1:
            # Se apenas um Excel foi gerado, envia diretamente
            info_arquivo_unico = arquivos_excel_processados_info[0]
            nome_original_sem_ext = os.path.splitext(info_arquivo_unico["original_name"])[0]
            nome_download_excel = f"{nome_original_sem_ext}_extracao.xlsx"
            try:
                return send_file(info_arquivo_unico["excel_path"], as_attachment=True, download_name=nome_download_excel)
            except Exception as e_send_single:
                print(f"Erro ao enviar arquivo Excel único: {e_send_single}")
                flash('Erro ao preparar arquivo Excel para download.', 'error')
                return redirect(request.url)
        else:
            # Se múltiplos Excels foram gerados, cria um arquivo ZIP
            id_zip_unico = str(uuid.uuid4())
            nome_arquivo_zip = f"extracao_multipla_{id_zip_unico}.zip"
            caminho_zip_saida = os.path.join(app.config['OUTPUT_FOLDER'], nome_arquivo_zip)
            g.files_to_remove.append(caminho_zip_saida) # Marca ZIP para remoção

            try:
                with zipfile.ZipFile(caminho_zip_saida, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for info in arquivos_excel_processados_info:
                        nome_original_sem_ext = os.path.splitext(info["original_name"])[0]
                        # Nome do arquivo dentro do ZIP
                        nome_arquivo_no_zip = f"{nome_original_sem_ext}_extracao.xlsx" 
                        zipf.write(info["excel_path"], arcname=nome_arquivo_no_zip)
                
                return send_file(caminho_zip_saida, as_attachment=True, download_name="planilhas_extraidas.zip")
            except Exception as e_zip:
                print(f"Erro ao criar ou enviar arquivo ZIP: {e_zip}")
                flash('Erro ao criar arquivo ZIP para download.', 'error')
                return redirect(request.url)

    # Para requisições GET, apenas renderiza o formulário
    return render_template_string(HTML_FORM)

# Bloco para executar a aplicação Flask
if __name__ == '__main__':
    print(f"Pasta de Uploads Temporários: {os.path.abspath(UPLOAD_FOLDER)}")
    print(f"Pasta de Saídas Temporárias: {os.path.abspath(OUTPUT_FOLDER)}")
    # debug=True é útil para desenvolvimento, mas desabilite em produção.
    # host='0.0.0.0' torna o servidor acessível na sua rede local (use o IP da sua máquina).
    app.run(host='0.0.0.0', port=5000, debug=True)

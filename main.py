print("\n[*] Iniciando Webdark - Sistema Unificado de Libras...\n")
import cv2
import csv
import os
import sys
import math
import time
import threading
import warnings
import numpy as np
from collections import deque, Counter
from PIL import Image, ImageDraw, ImageFont

# sklearn, mediapipe/cvzone e pandas levavam ~10s para importar e travavam a
# abertura. pandas foi removido (o CSV agora é lido com o módulo csv) e os
# outros dois são importados em threads (veja _carga mais abaixo), enquanto a
# câmera é escolhida.

# ============================================================================
# CHANGELOG DAS MELHORIAS (leia antes de usar)
# ----------------------------------------------------------------------------
# 1) MOVIMENTO POR MÃO: antes o vetor de movimento era calculado só a partir
#    do pulso da hands[0] e DUPLICADO para as duas mãos. Agora cada mão tem
#    seu próprio buffer de trajetória (buffer_movimento_*), então gestos com
#    2 mãos ficam corretos e cada mão tem seu movimento real.
#
# 2) FEATURES DE MOVIMENTO MAIS RICAS (8 por mão, antes eram 4 duplicadas):
#    dx, dy, retilinearidade (reto vs curvo), nº de mudanças de direção
#    (detecta "ganchos"/"cantos" como em J e Z), velocidade, aceleração,
#    ângulo principal e proporção da caixa delimitadora da trajetória.
#    Isso ajuda MUITO a diferenciar J/Z/X/H de letras estáticas parecidas,
#    porque antes só existia "deslocamento início->fim", que não captura a
#    curva característica do J ou o zigue-zague do Z.
#
# 3) NORMALIZAÇÃO POR ESCALA DA MÃO: o movimento agora é dividido pelo
#    tamanho da palma, então funciona igual perto ou longe da câmera.
#
# 4) CURL DOS DEDOS: 5 novas features estáticas (uma por dedo) medindo se o
#    dedo está esticado ou fechado, comparando a ponta com a articulação
#    intermediária. Ajuda a diferenciar letras com silhueta parecida mas
#    dedos em posições diferentes (ex.: M/N/S/T).
#
# 5) ORDENAÇÃO ESTÁVEL DAS MÃOS: as mãos agora são ordenadas por tipo
#    (Esquerda/Direita) antes de virar features, evitando que a posição no
#    vetor de features "troque de lugar" entre frames.
#
# 6) CAPTURA REAL EM VEZ DE SÓ RUÍDO SINTÉTICO: ao clicar em GRAVAR, o
#    sistema agora grava ~1 a 1.6s de frames REAIS da mão (parada com
#    micro-variação natural, ou o gesto de movimento completo), e só usa
#    ruído gaussiano sintético como complemento se poucas amostras reais
#    foram capturadas. Isso deixa o dataset mais realista.
#
# 7) TREINO MAIS ROBUSTO: agora usa train_test_split para estimar a
#    acurácia, class_weight='balanced' (ajuda quando um sinal tem menos
#    amostras que outro), e avisa quando algum sinal tem poucas amostras.
#
# 8) CORREÇÃO DE BUGS:
#    - Duas chamadas a cv2.waitKey por iteração do loop principal podiam
#      "roubar" teclas uma da outra. Agora a tecla lida em um frame só é
#      usada no próximo (padrão consistente), e o atalho de teclado [S]
#      não conflita mais com o clique do mouse.
#    - Variáveis 'predicao_estavel' e 'confianca' eram checadas via
#      `'predicao_estavel' in locals()`, o que é frágil. Agora são
#      inicializadas explicitamente no início do modo Escritor.
#    - Limiar de confiança para aceitar uma predição subiu de 30% para 45%,
#      e agora exige maioria consistente no buffer (>=3 de 5), reduzindo
#      "tremulação" de letras erradas.
#
# 9) NOVA INTERFACE: tema escuro com as cores da logo, logo.png no menu e nos
#    cabeçalhos, botões arredondados com hover, texto via PIL (acentos e "ç"
#    agora aparecem corretos; cv2.putText mostrava "?"), atalhos de teclado
#    (1/2/3 no menu, digitar a letra no painel), cursor piscando e quebra de
#    linha por largura real no Escritor, barras de precisão/confirmação.
#    Correções: CSV não é mais relido a cada frame; "Regravar" só apaga as
#    amostras antigas depois de gravar as novas; apagar pede confirmação;
#    botão Voltar disponível antes de escolher um sinal; falha de câmera não
#    fecha mais o programa; fechar a janela pelo "X" encerra corretamente.
# ============================================================================

warnings.filterwarnings('ignore', category=UserWarning)

# Troca de thread mais frequente (padrão 5 ms): com threads em segundo plano
# (detector de mãos, treino, voz), cada chamada do cv2 que libera o GIL podia
# esperar até 5 ms para voltar, derrubando o fps da interface.
sys.setswitchinterval(0.0005)

# --- SISTEMA DE VOZ (TEXT-TO-SPEECH OFFLINE NATIVO) ---
def falar_texto_async(texto):
    if not texto or not texto.strip():
        return
    def executar_fala():
        try:
            import win32com.client
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            speaker.Speak(texto)
        except Exception as e:
            print(f"[!] Erro ao reproduzir áudio: {e}")
    threading.Thread(target=executar_fala, daemon=True).start()

# --- CORRETOR DE PALAVRAS E VALIDADOR DE PORTUGUÊS ---
class CorretorLinguagem:
    def __init__(self):
        self.win_checker = None
        self.usar_win = False

        try:
            import win32com.client
            factory = win32com.client.Dispatch("SpellCheck.SpellCheckerFactory")
            self.win_checker = factory.CreateSpellChecker("pt-BR")
            self.usar_win = True
        except Exception:
            self.usar_win = False

        self.vocabulario = {
            "OI", "OLA", "TUDO", "BEM", "BOM", "DIA", "TARDE", "NOITE", "VOCE", "COMO",
            "VAI", "SIM", "NAO", "POR", "FAVOR", "OBRIGADO", "OBRIGADA", "MUITO", "AMIGO",
            "AMIGA", "CASA", "AMOR", "VIDA", "NOME", "MEU", "MINHA", "QUERO", "AJUDA",
            "SISTEMA", "LIBRAS", "SINAL", "A", "E", "O", "UM", "UMA", "DE", "DO", "DA",
            "EM", "NO", "NA", "QUE", "COM", "PARA", "SE", "MAIS", "MAS", "TEM", "SER",
            "ESTA", "ESTOU", "FAZER", "PODE", "AQUI", "HOJE", "AMANHA", "ONTEM", "AGORA"
        }

    def e_palavra_valida(self, palavra):
        if not palavra:
            return False

        p_clean = palavra.strip().upper()
        if len(p_clean) == 1 and p_clean in {"A", "E", "O", "U"}:
            return True
        if len(p_clean) < 2:
            return False

        if p_clean in self.vocabulario:
            return True

        if self.usar_win and self.win_checker:
            try:
                erros = self.win_checker.CheckSpelling(p_clean.lower())
                if erros.Next() is None:
                    return True
            except Exception:
                pass

        tem_vogal = any(c in "AEIOU" for c in p_clean)
        return len(p_clean) >= 2 and tem_vogal

corretor = CorretorLinguagem()

# --- FUNÇÕES DE DEGRADÊ E LOGS (WEBDARK) ---
def degrade_texto(texto, offset=0):
    resultado = ""
    comprimento = len(texto)
    if comprimento == 0: return ""
    for i in range(comprimento):
        t = (i + offset) / max(comprimento, 1)
        if t <= 0.5:
            progress = t * 2
            r = int(0 + (100 * progress))
            g = int(255 - (155 * progress))
            b = int(255 - (55 * progress))
        else:
            progress = (t - 0.5) * 2
            r = int(100 + (128 * progress))
            g = int(100 - (100 * progress))
            b = int(200 + (32 * progress))
        resultado += f'\033[38;2;{r};{g};{b}m{texto[i]}'
    return resultado + "\033[0m"

def degrade_linha2(texto, offset=0):
    resultado = ""
    comprimento = len(texto)
    if comprimento == 0: return ""
    for i in range(comprimento):
        t = (i + offset) / max(comprimento, 1)
        if t <= 0.5:
            progress = t * 2
            r = int(0 + (100 * progress))
            g = int(255 - (155 * progress))
            b = int(255 - (55 * progress))
        else:
            progress = (t - 0.5) * 2
            r = int(100 + (128 * progress))
            g = int(100 - (100 * progress))
            b = int(200 + (32 * progress))
        resultado += f'\033[38;2;0;0;0m\033[48;2;{r};{g};{b}m{texto[i]}'
    return resultado + "\033[0m"

def webdark():
    print(degrade_linha2("   ╔═══════════════════════════════════════════╗", offset=0))
    print(degrade_linha2("   ║ ▄▄▌ ▐ ▄▌▄▄▄ .▄▄▄▄· ·▄▄▄▄   ▄▄▄· ▄▄▄  ▄ •▄ ║", offset=0.2))
    print(degrade_linha2("   ║██· █▌▐█▀▄.▀·▐█ ▀█▪██▪ ██ ▐█ ▀█ ▀▄ █·█▌▄▌▪ ║", offset=0.4))
    print(degrade_linha2("   ║██▪▐█▐▐▌▐▀▀▪▄▐█▀▀█▄▐█· ▐█▌▄█▀▀█ ▐▀▀▄ ▐▀▀▄· ║", offset=0.6))
    print(degrade_linha2("   ║▐█▌██▐█▌▐█▄▄▌██▄▪▐███. ██ ▐█ ▪▐▌▐█•█▌▐█.█▌ ║", offset=0.8))
    print(degrade_linha2("   ║ ▀▀▀▀ ▀▪ ▀▀▀ ·▀▀▀▀ ▀▀▀▀▀•  ▀  ▀ .▀  ▀·▀  ▀ ║", offset=1.0))
    print(degrade_linha2("   ╚═══════════════════════════════════════════╝", offset=1.2))

# DirectShow abre a câmera quase instantaneamente no Windows; o backend padrão
# (MSMF) pode levar ~30s para inicializar.
BACKEND_CAMERA = cv2.CAP_DSHOW if os.name == 'nt' else cv2.CAP_ANY


ARQUIVO_CAMERA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "camera_padrao.txt")


# pygrabber (opcional) lista os NOMES dos dispositivos e abre câmeras que o
# OpenCV não consegue abrir, como o DroidCam. Instale com: pip install pygrabber
try:
    from pygrabber.dshow_graph import FilterGraph
except Exception:
    FilterGraph = None

nomes_cameras = {}


def nome_camera(indice):
    return nomes_cameras.get(indice, f"Câmera {indice}")


class CameraDirectShow:
    """Captura direto via DirectShow (pygrabber) com a mesma interface mínima do
    cv2.VideoCapture (isOpened/read/set/release)."""

    def __init__(self, indice):
        self._evento = threading.Event()
        self._frame = None
        self._graph = None
        self._aberta = False
        try:
            g = FilterGraph()
            g.add_video_input_device(indice)
            g.add_sample_grabber(self._ao_receber_frame)
            g.add_null_render()
            g.prepare_preview_graph()
            g.run()
            self._graph = g
            self._aberta = True
        except Exception as e:
            print(degrade_texto(f"[!] DirectShow não abriu a câmera {indice}: {e}", offset=0.8))
            self.release()

    def _ao_receber_frame(self, img):
        self._frame = np.ascontiguousarray(img)
        self._evento.set()

    def isOpened(self):
        return self._aberta

    def set(self, *args):
        return False

    def read(self):
        if not self._aberta:
            return False, None
        self._evento.clear()
        self._graph.grab_frame()
        if self._evento.wait(1.0):
            return True, self._frame
        return False, None

    def release(self):
        if self._graph is not None:
            try:
                self._graph.stop()
            except Exception:
                pass
        self._graph = None
        self._aberta = False


def _enumerar_cameras():
    """Devolve (índices, nomes). Com pygrabber lista por nome (instantâneo); sem
    ele testa os índices (aceita câmeras que abrem mesmo sem imagem ainda)."""
    nomes = {}
    if FilterGraph is not None:
        try:
            for i, n in enumerate(FilterGraph().get_input_devices()):
                nomes[i] = n.replace('\xa0', ' ')
        except Exception:
            nomes = {}
    if nomes:
        return sorted(nomes), nomes
    cams = []
    for i in range(8):
        cap_test = cv2.VideoCapture(i, BACKEND_CAMERA)
        if cap_test.isOpened():
            cams.append(i)
        cap_test.release()
    return cams, {}


def escanear_cameras():
    os.system('cls' if os.name == 'nt' else 'clear')
    webdark()
    print("\n" + degrade_texto("[*] Listando câmeras conectadas no sistema...", offset=0.1))

    cams_disponiveis, nomes = _enumerar_cameras()
    if not nomes:
        print(degrade_texto("[!] pygrabber não encontrado: sem nomes de câmera (pip install pygrabber).", offset=0.4))
    if not cams_disponiveis:
        print(degrade_texto("[!] Nenhuma câmera foi encontrada!", offset=0.5))
        exit()

    nomes_cameras.update(nomes)
    for i in cams_disponiveis:
        print(degrade_texto(f"[+] {i}: {nome_camera(i)}", offset=0.3))
    print(degrade_texto("[*] Troque de câmera pela interface (botão no topo ou tecla Tab).", offset=0.5))
    return cams_disponiveis


def atualizar_lista_cameras():
    """Relê os dispositivos (ex.: depois de abrir o DroidCam/plugar uma webcam)."""
    cams, nomes = _enumerar_cameras()
    if cams:
        cameras_disponiveis[:] = cams
        nomes_cameras.update(nomes)


def carregar_camera_padrao(cams):
    """Usa a última câmera escolhida na interface, se ainda estiver disponível."""
    try:
        with open(ARQUIVO_CAMERA, 'r', encoding='utf-8') as f:
            salva = int(f.read().strip())
        if salva in cams:
            return salva
    except (OSError, ValueError):
        pass
    return cams[0]


def salvar_camera_padrao(indice):
    try:
        with open(ARQUIVO_CAMERA, 'w', encoding='utf-8') as f:
            f.write(str(indice))
    except OSError:
        pass


# ============================================================================
# EXTRAÇÃO DE FEATURES (ESTÁTICAS + MOVIMENTO POR MÃO)
# ============================================================================
FEATURES_ESTATICAS_POR_MAO = 21 * 2 + 14 + 5   # pontos rotacionados + distâncias/ângulos + curl dos 5 dedos = 61
FEATURES_MOVIMENTO_POR_MAO = 8                  # dx, dy, retilinearidade, nº mudanças direção, vel, acel, ângulo, bbox
NUM_FEATURES = 2 * (FEATURES_ESTATICAS_POR_MAO + FEATURES_MOVIMENTO_POR_MAO)  # = 138


def ordenar_maos(hands):
    """Ordena as mãos de forma estável (Esquerda antes de Direita) para que a
    posição de cada mão no vetor de features não fique trocando entre frames."""
    if not hands:
        return hands
    return sorted(hands, key=lambda h: h.get('type', ''))


def calcular_features_movimento(buffer_pos):
    """
    buffer_pos: lista/deque de tuplas (x, y, escala_da_mao_no_frame).
    Retorna 8 features normalizadas pela escala média da mão, descrevendo o
    FORMATO da trajetória (não só o deslocamento início->fim), o que é
    essencial para diferenciar gestos como J (gancho) e Z (zigue-zague) de
    um simples deslocamento em linha reta.
    """
    if len(buffer_pos) < 4:
        return [0.0] * FEATURES_MOVIMENTO_POR_MAO

    escala_media = sum(p[2] for p in buffer_pos) / len(buffer_pos)
    escala_media = escala_media if escala_media > 1e-3 else 1.0

    xs = [p[0] for p in buffer_pos]
    ys = [p[1] for p in buffer_pos]

    dx = (xs[-1] - xs[0]) / escala_media
    dy = (ys[-1] - ys[0]) / escala_media
    deslocamento_liquido = math.hypot(dx, dy)

    comprimento_percurso = 0.0
    angulos = []
    for i in range(1, len(buffer_pos)):
        seg_dx = (xs[i] - xs[i - 1]) / escala_media
        seg_dy = (ys[i] - ys[i - 1]) / escala_media
        comprimento_percurso += math.hypot(seg_dx, seg_dy)
        if seg_dx != 0 or seg_dy != 0:
            angulos.append(math.atan2(seg_dy, seg_dx))

    comprimento_percurso = max(comprimento_percurso, 1e-4)
    # retilinearidade ~1.0 = movimento em linha reta; valores baixos indicam
    # trajetória curva/com cantos (típico de J, Z, X, H)
    retilinearidade = round(deslocamento_liquido / comprimento_percurso, 4)

    mudancas_direcao = 0
    for i in range(1, len(angulos)):
        diff = abs(angulos[i] - angulos[i - 1])
        if diff > math.pi:
            diff = 2 * math.pi - diff
        if diff > math.radians(35):
            mudancas_direcao += 1

    largura = (max(xs) - min(xs)) / escala_media
    altura = (max(ys) - min(ys)) / escala_media
    proporcao_bbox = round(largura / altura, 4) if altura > 1e-4 else 0.0

    metade = max(len(buffer_pos) // 2, 1)
    v1 = math.hypot(xs[metade] - xs[0], ys[metade] - ys[0]) / max(metade, 1)
    v2 = math.hypot(xs[-1] - xs[metade], ys[-1] - ys[metade]) / max(len(buffer_pos) - metade, 1)
    aceleracao = round((v2 - v1) / escala_media, 4)

    velocidade = round(comprimento_percurso / len(buffer_pos), 4)
    angulo_principal = round(math.atan2(dy, dx), 4)

    return [round(dx, 4), round(dy, 4), retilinearidade, float(mudancas_direcao),
            velocidade, aceleracao, angulo_principal, proporcao_bbox]


def atualizar_buffers_movimento(hands, buffers):
    """Atualiza (por índice de mão já ordenada) os buffers de trajetória do
    pulso, normalizados pela escala da palma naquele frame. Limpa o buffer da
    mão que sumiu do quadro para não misturar trajetórias descontínuas."""
    ocupados = set()
    for i in range(2):
        if i < len(hands):
            lm = hands[i]['lmList']
            wrist = lm[0]
            dx_palm = lm[9][0] - wrist[0]
            dy_palm = lm[9][1] - wrist[1]
            escala = math.hypot(dx_palm, dy_palm) or 1.0
            buffers[i].append((wrist[0], wrist[1], escala))
            ocupados.add(i)
    for i in range(2):
        if i not in ocupados:
            buffers[i].clear()


def features_movimento_das_maos(buffers):
    return [
        calcular_features_movimento(buffers[i]) if len(buffers[i]) >= 4 else [0.0] * FEATURES_MOVIMENTO_POR_MAO
        for i in range(2)
    ]


def extrair_features_avancadas(hands_data, movimento_por_mao=None):
    """
    hands_data: lista de mãos (formato cvzone), IDEALMENTE já ordenada com
                ordenar_maos() para manter consistência entre frames.
    movimento_por_mao: lista com até 2 listas de FEATURES_MOVIMENTO_POR_MAO
                        valores (uma por mão). Se None, usa zeros.
    """
    if movimento_por_mao is None:
        movimento_por_mao = [[0.0] * FEATURES_MOVIMENTO_POR_MAO, [0.0] * FEATURES_MOVIMENTO_POR_MAO]

    features_finais = []

    for idx_mao in range(2):
        if idx_mao < len(hands_data):
            hand = hands_data[idx_mao]
            lmList = hand['lmList']
            wrist_x, wrist_y = lmList[0][0], lmList[0][1]

            dx_palm = lmList[9][0] - wrist_x
            dy_palm = lmList[9][1] - wrist_y
            tamanho_palma = math.hypot(dx_palm, dy_palm)
            escala = tamanho_palma if tamanho_palma > 0 else 1.0

            # Orientação real da mão no espaço (diferencia H e P de U)
            angulo_orientacao_real = math.atan2(dy_palm, dx_palm)

            pontos_relativos = []
            for lm in lmList:
                pontos_relativos.append(((lm[0] - wrist_x) / escala, (lm[1] - wrist_y) / escala))

            # Rotação virtual para normalizar orientação da mão na hora de
            # comparar formato dos dedos
            x9, y9 = pontos_relativos[9]
            angulo_atual = math.atan2(y9, x9)
            angulo_alvo = -math.pi / 2.0
            delta_angulo = angulo_alvo - angulo_atual

            cos_a = math.cos(delta_angulo)
            sin_a = math.sin(delta_angulo)

            pontos_rotacionados = []
            for x, y in pontos_relativos:
                xr = x * cos_a - y * sin_a
                yr = x * sin_a + y * cos_a
                pontos_rotacionados.extend([round(xr, 4), round(yr, 4)])

            def dist(p1, p2):
                x1, y1 = pontos_rotacionados[p1 * 2], pontos_rotacionados[p1 * 2 + 1]
                x2, y2 = pontos_rotacionados[p2 * 2], pontos_rotacionados[p2 * 2 + 1]
                return round(math.hypot(x2 - x1, y2 - y1), 4)

            dist_4_8 = dist(4, 8)
            dist_4_12 = dist(4, 12)
            dist_4_16 = dist(4, 16)
            dist_4_20 = dist(4, 20)
            dist_8_12 = dist(8, 12)
            dist_12_16 = dist(12, 16)
            dist_16_20 = dist(16, 20)

            dist_8_0 = dist(8, 0)
            dist_12_0 = dist(12, 0)
            dist_16_0 = dist(16, 0)
            dist_20_0 = dist(20, 0)

            # Cruzamento de dedos (detecta R vs U/V)
            cruzamento_dedos = pontos_rotacionados[8 * 2] - pontos_rotacionados[12 * 2]

            # Abertura angular do V (diferencia U de V)
            v_indicador = (pontos_rotacionados[8 * 2] - pontos_rotacionados[5 * 2],
                           pontos_rotacionados[8 * 2 + 1] - pontos_rotacionados[5 * 2 + 1])
            v_medio = (pontos_rotacionados[12 * 2] - pontos_rotacionados[9 * 2],
                       pontos_rotacionados[12 * 2 + 1] - pontos_rotacionados[9 * 2 + 1])
            ang_ind = math.atan2(v_indicador[1], v_indicador[0])
            ang_med = math.atan2(v_medio[1], v_medio[0])
            abertura_u_v = round(ang_ind - ang_med, 4)

            # NOVO: curl (dobra) de cada dedo — compara a distância da ponta
            # ao pulso com a da articulação intermediária ao pulso. Valor
            # positivo = dedo esticado, negativo/perto de zero = dedo dobrado.
            # Ajuda a diferenciar letras com silhueta parecida mas dedos em
            # posições diferentes (ex.: M, N, S, T).
            curl_polegar = dist(4, 0) - dist(3, 0)
            curl_indicador = dist(8, 0) - dist(6, 0)
            curl_medio = dist(12, 0) - dist(10, 0)
            curl_anelar = dist(16, 0) - dist(14, 0)
            curl_minimo = dist(20, 0) - dist(18, 0)

            extra_features = [
                dist_4_8, dist_4_12, dist_4_16, dist_4_20,
                dist_8_12, dist_12_16, dist_16_20,
                dist_8_0, dist_12_0, dist_16_0, dist_20_0,
                round(angulo_orientacao_real, 4),
                round(cruzamento_dedos, 4),
                abertura_u_v,
                curl_polegar, curl_indicador, curl_medio, curl_anelar, curl_minimo,
            ]

            features_finais.extend(pontos_rotacionados + extra_features)  # 42 + 14 + 5 = 61
        else:
            features_finais.extend([0.0] * FEATURES_ESTATICAS_POR_MAO)

        mov = movimento_por_mao[idx_mao] if idx_mao < len(movimento_por_mao) else [0.0] * FEATURES_MOVIMENTO_POR_MAO
        features_finais.extend(mov)

    return features_finais

# ============================================================================
# GERENCIAMENTO DE BANCO DE DADOS
# ============================================================================
NOME_ARQUIVO = "banco_libras.csv"


def _escrever_cabecalho():
    headers = ['label'] + [f'f_{i}' for i in range(NUM_FEATURES)]
    with open(NOME_ARQUIVO, mode='w', newline='', encoding='utf-8') as f:
        csv.writer(f).writerow(headers)


def _ler_csv():
    """Devolve (cabecalho, linhas) do banco; linhas vazias são ignoradas."""
    with open(NOME_ARQUIVO, mode='r', newline='', encoding='utf-8') as f:
        leitor = csv.reader(f)
        cabecalho = next(leitor, [])
        return cabecalho, [r for r in leitor if r]


def inicializar_csv():
    if not os.path.exists(NOME_ARQUIVO) or os.stat(NOME_ARQUIVO).st_size == 0:
        _escrever_cabecalho()
        return
    try:
        with open(NOME_ARQUIVO, mode='r', newline='', encoding='utf-8') as f:
            cabecalho = next(csv.reader(f), [])
        if len(cabecalho) != NUM_FEATURES + 1:
            # Formato antigo/incompatível (ex.: veio da versão com 120
            # features) -> reseta para o novo formato de 138 features.
            _escrever_cabecalho()
    except Exception:
        _escrever_cabecalho()


def obter_contagem_labels():
    if not os.path.exists(NOME_ARQUIVO) or os.stat(NOME_ARQUIVO).st_size == 0:
        return {}
    try:
        _, linhas = _ler_csv()
        return dict(Counter(r[0] for r in linhas))
    except Exception:
        inicializar_csv()
        return {}


def apagar_amostras_label(label_alvo):
    if not os.path.exists(NOME_ARQUIVO) or os.stat(NOME_ARQUIVO).st_size == 0:
        return 0
    try:
        cabecalho, linhas = _ler_csv()
        mantidas = [r for r in linhas if r[0] != str(label_alvo)]
        removidos = len(linhas) - len(mantidas)
        with open(NOME_ARQUIVO, mode='w', newline='', encoding='utf-8') as f:
            escritor = csv.writer(f)
            escritor.writerow(cabecalho)
            escritor.writerows(mantidas)
        return removidos
    except Exception:
        return 0


# ============================================================================
# TEMA E COMPONENTES DE UI
# ============================================================================
LARGURA_CANVAS, ALTURA_CANVAS = 1280, 720
NOME_JANELA = "Webdark - Sistema Unificado de Libras"
CAMINHO_LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")

# Paleta em BGR (cores da logo: laranja + azul claro sobre fundo escuro)
COR_BG_TOPO = (34, 26, 22)
COR_BG_BASE = (20, 15, 13)
COR_CARD = (44, 36, 31)
COR_CARD_ALT = (58, 49, 43)
COR_BORDA = (86, 74, 66)
COR_TEXTO = (240, 236, 232)
COR_TEXTO2 = (185, 175, 165)
COR_MUDO = (128, 118, 110)
COR_ACENTO = (80, 155, 245)
COR_ACENTO2 = (225, 180, 135)
COR_OK = (130, 210, 100)
COR_ERRO = (100, 100, 235)
COR_AVISO = (60, 175, 255)

# estilo -> (fundo, fundo_hover, borda)
ESTILOS_BOTAO = {
    'primario': ((48, 98, 170), (64, 125, 210), COR_ACENTO),
    'ok':       ((52, 95, 40),  (68, 125, 52),  COR_OK),
    'aviso':    ((35, 88, 140), (45, 112, 178), COR_AVISO),
    'perigo':   ((50, 50, 120), (66, 66, 160),  COR_ERRO),
    'azul':     ((120, 85, 50), (150, 108, 65), COR_ACENTO2),
    'neutro':   (COR_CARD_ALT,  (74, 63, 56),   COR_BORDA),
}

_fila_texto = []
_cache_fontes = {}


def _caminho_fonte(negrito):
    pasta = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts')
    candidatas = ['segoeuib.ttf', 'arialbd.ttf'] if negrito else ['segoeui.ttf', 'arial.ttf']
    for nome in candidatas:
        caminho = os.path.join(pasta, nome)
        if os.path.exists(caminho):
            return caminho
    return None


def fonte(tam, negrito=False):
    chave = (tam, negrito)
    if chave not in _cache_fontes:
        caminho = _caminho_fonte(negrito)
        try:
            _cache_fontes[chave] = ImageFont.truetype(caminho, tam) if caminho else ImageFont.load_default()
        except Exception:
            _cache_fontes[chave] = ImageFont.load_default()
    return _cache_fontes[chave]


def txt(x, y, s, tam=18, cor=COR_TEXTO, negrito=False, ancora='l'):
    """Enfileira texto (com acentos) para ser desenhado em flush_texto().
    y é o centro vertical do texto; ancora: 'l' esquerda, 'm' centro, 'r' direita."""
    _fila_texto.append((x, y, str(s), tam, cor, negrito, ancora))


_cache_texto = {}
_bytes_cache_texto = 0


def _sprite_texto(s, tam, cor, negrito, ancora):
    """Texto já rasterizado e pré-multiplicado pela cor, cacheado. Rasterizar
    pelo PIL custava ~1 ms por string a cada frame. Devolve (inv, cor_x_alfa,
    deslocamento x, deslocamento y), onde inv = 255 - alfa."""
    global _bytes_cache_texto
    chave = (s, tam, cor, negrito, ancora)
    sprite = _cache_texto.get(chave)
    if sprite is None:
        f = fonte(tam, negrito)
        esq, topo, dir_, base = f.getbbox(s, anchor=ancora + 'm')
        im = Image.new('L', (max(1, dir_ - esq + 2), max(1, base - topo + 2)), 0)
        ImageDraw.Draw(im).text((1 - esq, 1 - topo), s, font=f, fill=255, anchor=ancora + 'm')
        alfa = np.asarray(im)
        alfa3 = np.repeat(alfa[:, :, None], 3, axis=2)
        inv = (255 - alfa3).astype(np.uint8)
        cor_alfa = (alfa3.astype(np.uint16) * np.array(cor, dtype=np.uint16))
        sprite = (np.ascontiguousarray(inv), np.ascontiguousarray(cor_alfa), esq - 1, topo - 1)
        _bytes_cache_texto += inv.nbytes + cor_alfa.nbytes
        if _bytes_cache_texto > 80_000_000:  # evita crescer sem limite com textos dinâmicos
            _cache_texto.clear()
            _bytes_cache_texto = inv.nbytes + cor_alfa.nbytes
        _cache_texto[chave] = sprite
    return sprite


def flush_texto(img):
    """Aplica todo o texto enfileirado sobre o canvas (cv2.putText não suporta
    acentos/cedilha), misturando só a região de cada string."""
    altura, largura = img.shape[:2]
    for x, y, s, tam, cor, negrito, ancora in _fila_texto:
        if not s:
            continue
        inv, cor_alfa, ox, oy = _sprite_texto(s, tam, tuple(cor), negrito, ancora)
        h, w = inv.shape[:2]
        x0, y0 = x + ox, y + oy
        cx0, cy0, cx1, cy1 = max(x0, 0), max(y0, 0), min(x0 + w, largura), min(y0 + h, altura)
        if cx0 >= cx1 or cy0 >= cy1:
            continue
        sl = (slice(cy0 - y0, cy1 - y0), slice(cx0 - x0, cx1 - x0))
        roi = img[cy0:cy1, cx0:cx1]
        # resultado = (fundo * (255 - alfa) + cor * alfa) / 255
        misto = cv2.add(cv2.multiply(roi, np.ascontiguousarray(inv[sl]), dtype=cv2.CV_16U),
                        np.ascontiguousarray(cor_alfa[sl]))
        roi[:] = cv2.convertScaleAbs(misto, alpha=1.0 / 255.0)
    _fila_texto.clear()


def largura_texto(s, tam, negrito=False):
    return fonte(tam, negrito).getlength(s)


def quebrar_linhas(texto, tam, largura, negrito=False):
    """Quebra texto em linhas pela largura real em pixels (quebra palavras
    maiores que a linha por caractere)."""
    f = fonte(tam, negrito)
    linhas, atual = [], ""
    for palavra in texto.split(' '):
        if not palavra:
            continue
        while f.getlength(palavra) > largura:
            corte = len(palavra)
            while corte > 1 and f.getlength(palavra[:corte]) > largura:
                corte -= 1
            if atual:
                linhas.append(atual)
                atual = ""
            linhas.append(palavra[:corte])
            palavra = palavra[corte:]
        teste = f"{atual} {palavra}" if atual else palavra
        if f.getlength(teste) <= largura:
            atual = teste
        else:
            linhas.append(atual)
            atual = palavra
    if atual or not linhas:
        linhas.append(atual)
    return linhas


def ret_arred(img, p1, p2, cor, r=10, esp=-1):
    """Retângulo com cantos arredondados e anti-aliasing (esp<0 = preenchido)."""
    x1, y1 = p1
    x2, y2 = p2
    r = max(0, min(r, (x2 - x1) // 2, (y2 - y1) // 2))
    if esp < 0:
        cv2.rectangle(img, (x1 + r, y1), (x2 - r, y2), cor, -1)
        cv2.rectangle(img, (x1, y1 + r), (x2, y2 - r), cor, -1)
        for cx, cy in ((x1 + r, y1 + r), (x2 - r, y1 + r), (x1 + r, y2 - r), (x2 - r, y2 - r)):
            cv2.circle(img, (cx, cy), r, cor, -1, cv2.LINE_AA)
    else:
        cv2.line(img, (x1 + r, y1), (x2 - r, y1), cor, esp, cv2.LINE_AA)
        cv2.line(img, (x1 + r, y2), (x2 - r, y2), cor, esp, cv2.LINE_AA)
        cv2.line(img, (x1, y1 + r), (x1, y2 - r), cor, esp, cv2.LINE_AA)
        cv2.line(img, (x2, y1 + r), (x2, y2 - r), cor, esp, cv2.LINE_AA)
        for (cx, cy), ang in (((x1 + r, y1 + r), 180), ((x2 - r, y1 + r), 270),
                              ((x2 - r, y2 - r), 0), ((x1 + r, y2 - r), 90)):
            cv2.ellipse(img, (cx, cy), (r, r), 0, ang, ang + 90, cor, esp, cv2.LINE_AA)


def cartao(img, x1, y1, x2, y2, fundo=COR_CARD, borda=COR_BORDA, r=12, esp=1):
    ret_arred(img, (x1, y1), (x2, y2), fundo, r)
    if borda is not None:
        ret_arred(img, (x1, y1), (x2, y2), borda, r, esp)


def barra(img, x1, y1, x2, y2, frac, cor):
    ret_arred(img, (x1, y1), (x2, y2), COR_CARD_ALT, (y2 - y1) // 2)
    frac = max(0.0, min(1.0, frac))
    if frac > 0:
        ret_arred(img, (x1, y1), (x1 + max(y2 - y1, int((x2 - x1) * frac)), y2), cor, (y2 - y1) // 2)


def botao(img, x1, y1, x2, y2, texto, estilo='neutro', tam=17, sub=None, atalho=None):
    """Desenha um botão e devolve True se o mouse está sobre ele."""
    hover = checar_hover(x1, y1, x2, y2, mouse_pos[0], mouse_pos[1])
    fundo, fundo_hover, borda = ESTILOS_BOTAO[estilo]
    ret_arred(img, (x1, y1), (x2, y2), fundo_hover if hover else fundo, 12)
    ret_arred(img, (x1, y1), (x2, y2), borda, 12, 2 if hover else 1)
    cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
    if sub:
        txt(cx, cy - 11, texto, tam, COR_TEXTO, True, 'm')
        txt(cx, cy + 14, sub, max(12, tam - 5), COR_TEXTO2, False, 'm')
    else:
        txt(cx, cy, texto, tam, COR_TEXTO, True, 'm')
    if atalho:
        txt(x2 - 16, cy, f"[{atalho}]", 13, COR_TEXTO2, False, 'r')
    return hover


def _gerar_fundo():
    t = np.linspace(0, 1, ALTURA_CANVAS, dtype=np.float32)[:, None, None]
    grad = np.array(COR_BG_TOPO, dtype=np.float32) * (1 - t) + np.array(COR_BG_BASE, dtype=np.float32) * t
    return np.repeat(grad, LARGURA_CANVAS, axis=1).astype(np.uint8)


FUNDO = _gerar_fundo()
_cache_logo = {}


def colar_logo(img, x, y, tam):
    """Recorta o emblema da logo e cola num quadrado arredondado claro."""
    if tam not in _cache_logo:
        pronto = None
        try:
            bruta = cv2.imread(CAMINHO_LOGO, cv2.IMREAD_UNCHANGED)
            if bruta is not None:
                if bruta.ndim == 3 and bruta.shape[2] == 4:
                    a = bruta[:, :, 3:4].astype(np.float32) / 255.0
                    base = np.full_like(bruta[:, :, :3], (244, 247, 247))
                    bruta = (bruta[:, :, :3] * a + base * (1 - a)).astype(np.uint8)
                h, w = bruta.shape[:2]
                meio = int(min(h, w) * 0.34)
                cx, cy = int(w * 0.505), int(h * 0.49)
                recorte = bruta[max(0, cy - meio):cy + meio, max(0, cx - meio):cx + meio]
                recorte = cv2.resize(recorte, (tam, tam), interpolation=cv2.INTER_AREA)
                mascara = np.zeros((tam, tam), dtype=np.uint8)
                ret_arred(mascara, (0, 0), (tam - 1, tam - 1), 255, max(6, tam // 5))
                pronto = (recorte, mascara[:, :, None].astype(np.float32) / 255.0)
        except Exception:
            pronto = None
        _cache_logo[tam] = pronto
    logo = _cache_logo[tam]
    if logo is None:
        return
    recorte, m = logo
    roi = img[y:y + tam, x:x + tam]
    roi[:] = (recorte * m + roi * (1 - m)).astype(np.uint8)


def botao_camera(img, x1, y1, x2, y2):
    """Botão que mostra a câmera atual e troca para a próxima ao clicar."""
    rotulo = f"CÂMERA {cam_index}: {nome_camera(cam_index)}"
    f = fonte(14, True)
    while len(rotulo) > 4 and f.getlength(rotulo) > (x2 - x1) - 130:
        rotulo = rotulo[:-2].rstrip() + "…"
    return botao(img, x1, y1, x2, y2, rotulo, 'neutro', 14, atalho="Tab")


def cabecalho(img, titulo, direita=None):
    """Desenha o cabeçalho e devolve True se o mouse está no botão de câmera."""
    cartao(img, 30, 20, 1250, 84)
    colar_logo(img, 38, 28, 48)
    txt(104, 52, titulo, 21, COR_TEXTO, True)
    if direita:
        txt(840, 52, direita, 15, COR_TEXTO2, False, 'r')
    return botao_camera(img, 860, 32, 1230, 72)


def rodape(img):
    cartao(img, 30, 650, 1250, 700, COR_BG_BASE, COR_BORDA)
    cv2.circle(img, (56, 675), 6, status_cor, -1, cv2.LINE_AA)
    txt(74, 675, status_msg, 16, COR_TEXTO)
    txt(1230, 675, "ESC: voltar", 13, COR_MUDO, False, 'r')


def colar_video(img, frame, x1, y1, x2, y2, cor_borda):
    img[y1:y2, x1:x2] = cv2.resize(frame, (x2 - x1, y2 - y1))
    cv2.rectangle(img, (x1, y1), (x2, y2), cor_borda, 2)


def _camera_funciona(camera):
    return camera is not None and camera.isOpened() and any(camera.read()[0] for _ in range(5))


def abrir_camera(indice=None):
    """Abre a câmera pelo OpenCV e, se falhar, direto pelo DirectShow (abre o
    DroidCam, que o OpenCV não consegue). Devolve None se nada funcionar."""
    indice = cam_index if indice is None else indice

    def via_opencv():
        camera = cv2.VideoCapture(indice, BACKEND_CAMERA)
        if camera.isOpened():
            camera.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
            camera.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        return camera

    def via_directshow():
        return CameraDirectShow(indice) if FilterGraph is not None else None

    tentativas = [via_directshow, via_opencv] if 'droidcam' in nome_camera(indice).lower() else [via_opencv, via_directshow]
    for abrir in tentativas:
        camera = abrir()
        if _camera_funciona(camera):
            return camera
        if camera is not None:
            camera.release()
    return None


def trocar_camera(novo_indice):
    """Abre a nova câmera e só então fecha a antiga (se falhar, mantém a atual)."""
    global cap, cam_index
    nova = abrir_camera(novo_indice)
    if nova is None:
        return False
    antiga = cap
    cap, cam_index = nova, novo_indice
    if antiga is not None:
        antiga.release()
    salvar_camera_padrao(novo_indice)
    for b in buffer_movimento_calibracao + buffer_movimento_escritor:
        b.clear()
    buffer_predicoes.clear()
    return True


def desenhar_texto_multilinhas(img, texto, x, y, largura=740, tam=32, passo=46, max_linhas=9):
    """Escreve o texto com quebra por largura real e cursor piscando."""
    linhas = quebrar_linhas(texto, tam, largura)[-max_linhas:]
    for i, linha in enumerate(linhas):
        txt(x, y + i * passo, linha, tam, COR_TEXTO)
    if int(time.time() * 2) % 2 == 0:
        ultima = linhas[-1] if linhas else ""
        cx = x + int(largura_texto(ultima, tam)) + (int(largura_texto(' ', tam)) if texto.endswith(' ') else 0)
        cy = y + (len(linhas) - 1) * passo
        cv2.line(img, (cx + 3, cy - tam // 2 - 2), (cx + 3, cy + tam // 2 + 2), COR_ACENTO, 2, cv2.LINE_AA)


# ============================================================================
# ESTADOS, ENTRADA E TREINO
# ============================================================================
STATE_MENU = 0
STATE_CALIB_GRID = 1
STATE_CALIB_ACTION = 2
STATE_CALIB_CAM = 3
STATE_WRITER_CAM = 4

LABELS = [chr(i) for i in range(ord('A'), ord('Z') + 1)] + [str(i) for i in range(10)]
LABELS_MOVIMENTO = {'J', 'X', 'H', 'Z'}

estado_atual = STATE_MENU
label_selecionada = None
modo_camera_acao = None
tipo_captura_movimento = False
status_msg = "Bem-vindo ao Sistema Webdark LIBRAS!"
status_cor = COR_ACENTO

mouse_pos = (0, 0)
mouse_clicked = False


# Geometria atual da imagem dentro da janela (para letterbox e mouse)
_view = {'escala': 1.0, 'ox': 0, 'oy': 0}


def apresentar(img):
    """Mostra o canvas 1280x720 ajustado ao tamanho atual da janela, mantendo a
    proporção (barras laterais/superiores se necessário)."""
    try:
        _, _, w, h = cv2.getWindowImageRect(NOME_JANELA)
    except cv2.error:
        w, h = LARGURA_CANVAS, ALTURA_CANVAS
    if w <= 0 or h <= 0:
        w, h = LARGURA_CANVAS, ALTURA_CANVAS
    escala = min(w / LARGURA_CANVAS, h / ALTURA_CANVAS)
    nw, nh = max(1, int(LARGURA_CANVAS * escala)), max(1, int(ALTURA_CANVAS * escala))
    if (nw, nh) == (LARGURA_CANVAS, ALTURA_CANVAS):
        redim = img
    else:
        interp = cv2.INTER_AREA if escala < 1 else cv2.INTER_LINEAR
        redim = cv2.resize(img, (nw, nh), interpolation=interp)
    if (nw, nh) == (w, h):
        quadro, ox, oy = redim, 0, 0
    else:
        quadro = np.empty((h, w, 3), dtype=np.uint8)
        quadro[:] = COR_BG_BASE
        ox, oy = (w - nw) // 2, (h - nh) // 2
        quadro[oy:oy + nh, ox:ox + nw] = redim
    _view.update(escala=escala, ox=ox, oy=oy)
    cv2.imshow(NOME_JANELA, quadro)


def gerenciar_clique_mouse(event, x, y, flags, param):
    global mouse_pos, mouse_clicked
    # converte coordenadas da janela para coordenadas do canvas 1280x720
    mouse_pos = (int((x - _view['ox']) / _view['escala']), int((y - _view['oy']) / _view['escala']))
    if event == cv2.EVENT_LBUTTONDOWN:
        mouse_clicked = True


def checar_hover(x1, y1, x2, y2, mx, my):
    return x1 <= mx <= x2 and y1 <= my <= y2


def treinar_modelo():
    """Valida o CSV e treina o Random Forest. Retorna (modelo|None, msg, cor)."""
    if not os.path.exists(NOME_ARQUIVO) or os.stat(NOME_ARQUIVO).st_size == 0:
        return None, "O banco de dados está vazio! Adicione sinais na calibração primeiro.", COR_ERRO
    try:
        cabecalho, linhas = _ler_csv()
        if len(cabecalho) - 1 != NUM_FEATURES:
            inicializar_csv()
            return None, "CSV incompatível! O arquivo foi resetado para a versão nova.", COR_ERRO
        if len(linhas) < 5:
            return None, "Poucas amostras salvas. Cadastre mais sinais na calibração!", COR_ERRO

        _thread_sklearn.join()
        RandomForestClassifier, train_test_split, accuracy_score = _carga['sk']

        y = np.array([r[0] for r in linhas])
        X = np.array([r[1:] for r in linhas], dtype=np.float64)
        contagem_classes = Counter(y.tolist())
        classes_poucas = [c for c, n in contagem_classes.items() if n < 5]
        if classes_poucas:
            print(degrade_texto(
                f"[!] Aviso: sinais com poucas amostras: {classes_poucas}. "
                f"Recomenda-se coletar mais para melhor precisão.", offset=0.5))

        print(degrade_texto("[*] Treinando modelo Random Forest...", offset=0.3))

        pode_validar = min(contagem_classes.values()) >= 2 and len(linhas) >= 20
        if pode_validar:
            X_train, X_test, y_train, y_test = train_test_split(
                X, y, test_size=0.15, random_state=42, stratify=y)
        else:
            X_train, y_train, X_test, y_test = X, y, None, None

        modelo = RandomForestClassifier(
            n_estimators=150, max_depth=20, min_samples_leaf=2,
            class_weight='balanced', random_state=42, n_jobs=-1)
        modelo.fit(X_train, y_train)

        if X_test is not None:
            acc = accuracy_score(y_test, modelo.predict(X_test))
            print(degrade_texto(f"[+] Acurácia estimada (conjunto de teste): {acc*100:.1f}%", offset=0.4))
            return modelo, f"Modelo treinado! Acurácia estimada: {acc*100:.1f}%", COR_OK
        return modelo, "Modelo treinado (poucas amostras para validação separada).", COR_OK
    except Exception as e:
        return None, f"Erro ao treinar modelo: {e}", COR_ERRO


class FlorestaRapida:
    """Inferência de 1 amostra do RandomForest percorrendo todas as árvores de
    uma vez com numpy. O predict_proba do sklearn faz um laço Python por árvore
    (~50 ms para 150 árvores) e segura o GIL, travando o resto do programa."""

    def __init__(self, rf):
        arvores = [e.tree_ for e in rf.estimators_]
        n, k = len(arvores), len(rf.classes_)
        maxn = max(t.node_count for t in arvores)
        self.esq = np.zeros((n, maxn), dtype=np.int32)
        self.dir = np.zeros((n, maxn), dtype=np.int32)
        self.feat = np.zeros((n, maxn), dtype=np.int32)
        self.lim = np.zeros((n, maxn), dtype=np.float64)
        self.prob = np.zeros((n, maxn, k), dtype=np.float64)
        for i, t in enumerate(arvores):
            m = t.node_count
            nos = np.arange(m, dtype=np.int32)
            folha = t.children_left == -1
            self.esq[i, :m] = np.where(folha, nos, t.children_left)   # folha aponta para si mesma
            self.dir[i, :m] = np.where(folha, nos, t.children_right)
            self.feat[i, :m] = np.where(folha, 0, t.feature)
            self.lim[i, :m] = t.threshold
            valor = t.value[:, 0, :]
            self.prob[i, :m] = valor / valor.sum(axis=1, keepdims=True)
        self.profundidade = max(t.max_depth for t in arvores)
        self.classes = [str(c) for c in rf.classes_]
        self._arvores = np.arange(n)

    def predict_proba(self, x):
        x = np.asarray(x, dtype=np.float32)
        no = np.zeros(len(self._arvores), dtype=np.int32)
        for _ in range(self.profundidade):
            f = self.feat[self._arvores, no]
            vai_esq = x[f] <= self.lim[self._arvores, no]
            no = np.where(vai_esq, self.esq[self._arvores, no], self.dir[self._arvores, no])
        return self.prob[self._arvores, no].mean(axis=0)

    def prever(self, x):
        p = self.predict_proba(x)
        i = int(np.argmax(p))
        return self.classes[i], float(p[i] * 100)


class DetectorAssincrono:
    """Roda o detector de mãos (~60 ms por frame, mesmo sem mão na imagem) numa
    thread própria. O loop principal continua mostrando o vídeo na velocidade
    da câmera e só lê o último resultado pronto. O lock garante que o detector
    nunca seja usado por duas threads ao mesmo tempo."""

    def __init__(self, detector):
        self.detector = detector
        self._lock = threading.Lock()
        self._evento = threading.Event()
        self._entrada = None         # (id, frame)
        self._id_entrada = 0
        self._id_processado = 0
        self._resultado = (0, [])    # (id, mãos)
        self._ativo = True
        threading.Thread(target=self._loop, daemon=True).start()

    def enviar(self, frame):
        self._id_entrada += 1
        self._entrada = (self._id_entrada, frame.copy())  # cópia: o loop principal desenha no frame original
        self._evento.set()

    def ultimo(self):
        return self._resultado

    def detectar_sync(self, frame):
        """Detecção síncrona (usada na gravação de amostras, para manter a mesma
        taxa de quadros com que o dataset foi gravado)."""
        with self._lock:
            return self.detector.findHands(frame, draw=False)[0]

    def _loop(self):
        while self._ativo:
            self._evento.wait(0.5)
            self._evento.clear()
            entrada = self._entrada
            if entrada is None or entrada[0] == self._id_processado:
                continue
            try:
                with self._lock:
                    maos = self.detector.findHands(entrada[1], draw=False)[0]
                self._resultado = (entrada[0], maos)
            except Exception as e:
                print(degrade_texto(f"[!] Erro na detecção: {e}", offset=0.8))
            self._id_processado = entrada[0]


CONEXOES_MAO = [(0, 1), (1, 2), (2, 3), (3, 4), (0, 5), (5, 6), (6, 7), (7, 8), (5, 9), (9, 10), (10, 11),
                (11, 12), (9, 13), (13, 14), (14, 15), (15, 16), (13, 17), (17, 18), (18, 19), (19, 20), (0, 17)]


def desenhar_maos(img, maos):
    """Desenha o esqueleto de cada mão detectada (substitui o draw=True do cvzone)."""
    for mao in maos:
        pts = [(p[0], p[1]) for p in mao['lmList']]
        for a, b in CONEXOES_MAO:
            cv2.line(img, pts[a], pts[b], COR_ACENTO2, 2, cv2.LINE_AA)
        for p in pts:
            cv2.circle(img, p, 4, COR_ACENTO, -1, cv2.LINE_AA)


# --- CACHE E TREINO EM SEGUNDO PLANO ---
# O modelo só é retreinado quando o CSV muda; depois de cada gravação/exclusão
# o treino já começa em segundo plano, então entrar no Escritor é instantâneo.
_treino = {'assinatura': None, 'modelo': None, 'msg': '', 'cor': COR_ACENTO, 'thread': None}


def assinatura_csv():
    try:
        st = os.stat(NOME_ARQUIVO)
        return (st.st_mtime_ns, st.st_size)
    except OSError:
        return None


def _treino_worker():
    assinatura = assinatura_csv()
    modelo, msg, cor = treinar_modelo()
    _treino.update(assinatura=assinatura, modelo=modelo, msg=msg, cor=cor)


def iniciar_treino_background():
    t = _treino['thread']
    if t is None or not t.is_alive():
        _treino['thread'] = threading.Thread(target=_treino_worker, daemon=True)
        _treino['thread'].start()


def modelo_esta_pronto():
    t = _treino['thread']
    assinatura = assinatura_csv()
    return (t is None or not t.is_alive()) and assinatura is not None and _treino['assinatura'] == assinatura


def obter_modelo():
    """Devolve (modelo|None, msg, cor), reaproveitando o treino em cache."""
    t = _treino['thread']
    if t is not None and t.is_alive():
        t.join()
    assinatura = assinatura_csv()
    if assinatura is None or _treino['assinatura'] != assinatura:
        _treino_worker()
    return _treino['modelo'], _treino['msg'], _treino['cor']


# --- INICIALIZAÇÃO COMPLETA ---
# Imports pesados em paralelo com a escolha da câmera (que também demora).
_carga = {}


def _carregar_mediapipe():
    from cvzone.HandTrackingModule import HandDetector
    _carga['HandDetector'] = HandDetector


def _carregar_sklearn():
    from sklearn.ensemble import RandomForestClassifier
    from sklearn.model_selection import train_test_split
    from sklearn.metrics import accuracy_score
    _carga['sk'] = (RandomForestClassifier, train_test_split, accuracy_score)


_thread_mediapipe = threading.Thread(target=_carregar_mediapipe, daemon=True)
_thread_sklearn = threading.Thread(target=_carregar_sklearn, daemon=True)
_thread_mediapipe.start()
_thread_sklearn.start()

inicializar_csv()
cameras_disponiveis = escanear_cameras()
cam_index = carregar_camera_padrao(cameras_disponiveis)

if _thread_mediapipe.is_alive():
    print(degrade_texto("[*] Carregando módulos de reconhecimento de mãos...", offset=0.4))
_thread_mediapipe.join()
if 'HandDetector' not in _carga:
    raise RuntimeError("Falha ao importar cvzone/mediapipe (veja o erro acima).")
detector = _carga['HandDetector'](detectionCon=0.8, maxHands=2)
detector_async = DetectorAssincrono(detector)

# A câmera fica aberta durante toda a sessão (sem reabrir a cada tela). Tenta a
# última usada primeiro e, se não abrir, as demais.
cap = None
for _cand in [cam_index] + [c for c in cameras_disponiveis if c != cam_index]:
    print(degrade_texto(f"[*] Abrindo câmera {_cand}: {nome_camera(_cand)}...", offset=0.5))
    cap = abrir_camera(_cand)
    if cap is not None:
        cam_index = _cand
        break
    print(degrade_texto(f"[!] Não foi possível abrir a câmera {_cand}.", offset=0.8))
if cap is None:
    print(degrade_texto("[!] Nenhuma câmera funcionou. Se for DroidCam, conecte o cliente ao celular.", offset=0.8))
    exit()
iniciar_treino_background()  # espera o sklearn terminar de importar dentro da thread

cv2.namedWindow(NOME_JANELA, cv2.WINDOW_NORMAL)
cv2.resizeWindow(NOME_JANELA, LARGURA_CANVAS, ALTURA_CANVAS)
cv2.setMouseCallback(NOME_JANELA, gerenciar_clique_mouse)

# Variáveis do Escritor
modelo_rf = None
floresta = None            # FlorestaRapida do modelo atual
rid_buffer_escritor = -1   # id da última detecção de mãos já usada nos buffers de trajetória
buffer_predicoes = deque(maxlen=5)
MIN_VOTOS_ESTAVEL = 3          # de 5 amostras no buffer, exige maioria mínima
CONFIDENCE_THRESHOLD = 45.0    # antes era 30 — muito permissivo, gerava letras erradas

# Buffers de trajetória (um par de deques por modo, um slot por mão já ordenada)
buffer_movimento_calibracao = [deque(maxlen=20), deque(maxlen=20)]
buffer_movimento_escritor = [deque(maxlen=20), deque(maxlen=20)]

texto_construido = ""
ultima_letra_detectada = None
tempo_inicio_hold = 0.0
letra_adicionada = False
TEMPO_HOLD_NECESSARIO = 0.9

auto_espaco_ativo = True
tempo_ultima_letra_adicionada = 0.0
espaco_pendente = False
TEMPO_AUTO_ESPACO_VALIDO = 1.0
TEMPO_AUTO_ESPACO_INVALIDO = 3.5

# Predição corrente do modo escritor (inicializadas aqui para nunca depender
# de checagem frágil via locals())
predicao_estavel = "Identificando..."
confianca = 0.0

# Cache da contagem de amostras (antes o CSV era relido a cada frame!)
counts = {}
recarregar_counts = True

# Confirmação em 2 cliques para apagar amostras
confirmar_apagar_ate = 0.0

# Seletor de câmera (janela modal aberta pelo botão do topo ou pela tecla Tab)
modal_camera = False
camera_escolhida = None

# --- LOOP PRINCIPAL DO SISTEMA ---
tecla_atual = 255  # última tecla lida (usada no início da iteração seguinte,
                    # para não competir com o cv2.waitKey do final do loop)

while True:
    canvas = FUNDO.copy()
    h_cam = False
    clique_modal = False
    if modal_camera:  # com o seletor aberto, a tela de baixo não recebe cliques/teclas
        clique_modal, mouse_clicked, tecla_atual = mouse_clicked, False, 255
    if recarregar_counts:
        counts = obter_contagem_labels()
        recarregar_counts = False
    total_amostras = sum(counts.values())
    total_sinais = len(counts)

    # ==========================================
    # ESTADO 0: TELA PRINCIPAL (SELEÇÃO INICIAL)
    # ==========================================
    if estado_atual == STATE_MENU:
        colar_logo(canvas, (LARGURA_CANVAS - 150) // 2, 28, 150)
        txt(LARGURA_CANVAS // 2, 222, "SISTEMA DE RECONHECIMENTO DE LIBRAS", 30, COR_TEXTO, True, 'm')
        txt(LARGURA_CANVAS // 2, 254, "Calibre os sinais, treine o modelo e escreva com as mãos.", 16, COR_TEXTO2, False, 'm')

        h_cam = botao_camera(canvas, 860, 32, 1230, 72)
        h1 = botao(canvas, 320, 288, 960, 364, "MODO CALIBRAÇÃO / COLETOR DE SINAIS", 'ok', 19,
                   sub="Grave e gerencie as amostras de cada letra e número", atalho="1")
        h2 = botao(canvas, 320, 380, 960, 456, "MODO ESCRITOR / IDENTIFICADOR", 'azul', 19,
                   sub="Treina o modelo e transforma sinais em texto em tempo real", atalho="2")
        h3 = botao(canvas, 320, 472, 960, 548, "SAIR DO SISTEMA", 'perigo', 19, atalho="3")

        cartao(canvas, 320, 572, 960, 624, COR_CARD, COR_BORDA)
        txt(640, 598, f"Banco de dados:  {total_amostras} amostras  ·  {total_sinais}/{len(LABELS)} sinais calibrados",
            15, COR_TEXTO2, False, 'm')

        escolha_menu = 0
        if mouse_clicked:
            escolha_menu = 1 if h1 else 2 if h2 else 3 if h3 else 0
        elif tecla_atual in (ord('1'), ord('2'), ord('3')):
            escolha_menu = tecla_atual - ord('0')

        if escolha_menu == 1:
            estado_atual = STATE_CALIB_GRID
            status_msg = "Selecione uma letra ou número (clique ou digite a tecla)."
            status_cor = COR_ACENTO
        elif escolha_menu == 2:
            if not modelo_esta_pronto():  # só mostra aviso se realmente precisar treinar
                cartao(canvas, 440, 300, 840, 400, COR_CARD, COR_ACENTO, 14, 2)
                txt(640, 335, "Treinando modelo...", 22, COR_TEXTO, True, 'm')
                txt(640, 368, "Isso pode levar alguns segundos", 14, COR_TEXTO2, False, 'm')
                flush_texto(canvas)
                apresentar(canvas)
                cv2.waitKey(1)

            novo_modelo, status_msg, status_cor = obter_modelo()
            if novo_modelo is not None:
                modelo_rf = novo_modelo
                floresta = FlorestaRapida(modelo_rf)
                rid_buffer_escritor = -1
                buffer_movimento_escritor[0].clear()
                buffer_movimento_escritor[1].clear()
                buffer_predicoes.clear()
                predicao_estavel = "Identificando..."
                confianca = 0.0
                estado_atual = STATE_WRITER_CAM
        elif escolha_menu == 3:
            break

    # ==========================================
    # ESTADO 1 E 2: GRID E ACOES DA CALIBRAÇÃO
    # ==========================================
    elif estado_atual in (STATE_CALIB_GRID, STATE_CALIB_ACTION):
        h_cam = cabecalho(canvas, "PAINEL DE CALIBRAÇÃO",
                  f"{total_sinais}/{len(LABELS)} sinais  ·  {total_amostras} amostras")

        start_x, start_y = 50, 104
        box_w, box_h = 100, 75
        gap_x, gap_y = 15, 12
        label_clicada = None

        for idx, lbl in enumerate(LABELS):
            col = idx % 6
            row = idx // 6
            x1 = start_x + col * (box_w + gap_x)
            y1 = start_y + row * (box_h + gap_y)
            x2, y2 = x1 + box_w, y1 + box_h

            qtd = counts.get(str(lbl), 0)
            is_hover = checar_hover(x1, y1, x2, y2, mouse_pos[0], mouse_pos[1])
            selecionado = estado_atual == STATE_CALIB_ACTION and lbl == label_selecionada

            if selecionado:
                bg_color, border_color = (70, 58, 40), COR_ACENTO
            elif qtd > 0:
                bg_color = (52, 80, 46) if is_hover else (40, 62, 36)
                border_color = COR_OK
            else:
                bg_color = (72, 62, 56) if is_hover else COR_CARD
                border_color = COR_BORDA

            ret_arred(canvas, (x1, y1), (x2, y2), bg_color, 12)
            ret_arred(canvas, (x1, y1), (x2, y2), border_color, 12, 2 if (selecionado or is_hover) else 1)

            txt((x1 + x2) // 2, y1 + 33, lbl, 32, COR_TEXTO, True, 'm')
            txt((x1 + x2) // 2, y2 - 13, f"{qtd}x", 12,
                (COR_AVISO if 0 < qtd < 5 else COR_TEXTO2 if qtd else COR_MUDO), False, 'm')
            if lbl in LABELS_MOVIMENTO:
                txt(x2 - 8, y1 + 12, "MOV", 10, COR_AVISO, True, 'r')

            if mouse_clicked and is_hover:
                label_clicada = str(lbl)

        # Atalho de teclado: digitar o caractere seleciona o sinal
        if label_clicada is None and 32 < tecla_atual < 128 and chr(tecla_atual).upper() in LABELS:
            label_clicada = chr(tecla_atual).upper()

        if label_clicada is not None:
            label_selecionada = label_clicada
            tipo_captura_movimento = label_selecionada in LABELS_MOVIMENTO
            estado_atual = STATE_CALIB_ACTION
            confirmar_apagar_ate = 0.0
            status_msg = f"Caractere '{label_selecionada}' selecionado."
            status_cor = COR_ACENTO

        # Painel de ações
        cartao(canvas, 760, 104, 1250, 614)
        txt(780, 128, "AÇÕES DISPONÍVEIS", 13, COR_MUDO, True)

        if estado_atual == STATE_CALIB_ACTION:
            qtd_atual = counts.get(str(label_selecionada), 0)

            cartao(canvas, 780, 146, 1230, 226, COR_CARD_ALT, COR_ACENTO, 12, 2)
            cartao(canvas, 792, 158, 862, 214, COR_CARD, COR_BORDA, 10)
            txt(827, 187, label_selecionada, 36, COR_TEXTO, True, 'm')
            txt(882, 173, "Amostras salvas", 13, COR_TEXTO2)
            txt(882, 199, str(qtd_atual), 24, COR_AVISO if 0 < qtd_atual < 5 else COR_TEXTO, True)
            if 0 < qtd_atual < 5:
                txt(1216, 199, "poucas amostras", 12, COR_AVISO, False, 'r')

            ht = botao(canvas, 780, 242, 1230, 286,
                       "Tipo: DINÂMICO (movimento)" if tipo_captura_movimento else "Tipo: ESTÁTICO (sinal parado)",
                       'aviso' if tipo_captura_movimento else 'neutro', 15, atalho="clique")

            h_add = botao(canvas, 780, 304, 1230, 354, "+ ADICIONAR MAPEAMENTO", 'ok', 16)
            h_edit = botao(canvas, 780, 366, 1230, 416, "REGRAVAR (SOBRESCREVER)", 'aviso', 16)

            aguardando_confirmacao = time.time() < confirmar_apagar_ate
            h_del = botao(canvas, 780, 428, 1230, 478,
                          "CLIQUE DE NOVO PARA CONFIRMAR" if aguardando_confirmacao else "APAGAR AMOSTRAS",
                          'perigo', 16)
        else:
            ht = h_add = h_edit = h_del = False
            txt(1005, 300, "Selecione um sinal", 22, COR_TEXTO, True, 'm')
            txt(1005, 334, "Clique em uma caixa ou digite a tecla", 14, COR_TEXTO2, False, 'm')
            txt(1005, 356, "correspondente no teclado.", 14, COR_TEXTO2, False, 'm')
            cv2.rectangle(canvas, (800, 420), (816, 436), COR_OK, -1)
            txt(826, 428, "Sinal já calibrado", 13, COR_TEXTO2)
            txt(960, 428, "MOV", 11, COR_AVISO, True)
            txt(996, 428, "Sinal com movimento (J, X, H, Z)", 13, COR_TEXTO2)
            txt(800, 458, "0x–4x em laranja = poucas amostras (ideal: 30+)", 13, COR_TEXTO2)

        h_back = botao(canvas, 780, 540, 1230, 590, "← VOLTAR AO MENU PRINCIPAL", 'neutro', 15)

        if mouse_clicked and estado_atual == STATE_CALIB_ACTION:
            if ht:
                tipo_captura_movimento = not tipo_captura_movimento
            elif h_add or h_edit:
                modo_camera_acao = 'ADICIONAR' if h_add else 'EDITAR'
                estado_atual = STATE_CALIB_CAM
                buffer_movimento_calibracao[0].clear()
                buffer_movimento_calibracao[1].clear()
                if h_add:
                    status_msg = f"Posicione a mão para '{label_selecionada}' e clique em GRAVAR."
                    status_cor = COR_OK
                else:
                    status_msg = f"As amostras antigas só serão trocadas se você gravar '{label_selecionada}'."
                    status_cor = COR_AVISO
            elif h_del:
                if time.time() < confirmar_apagar_ate:
                    removidos = apagar_amostras_label(label_selecionada)
                    recarregar_counts = True
                    iniciar_treino_background()
                    confirmar_apagar_ate = 0.0
                    status_msg = f"Removidas {removidos} amostras de '{label_selecionada}'."
                    status_cor = COR_ERRO
                    estado_atual = STATE_CALIB_GRID
                else:
                    confirmar_apagar_ate = time.time() + 3.0
                    status_msg = f"Isso apaga todas as amostras de '{label_selecionada}'. Clique de novo para confirmar."
                    status_cor = COR_AVISO
        if mouse_clicked and h_back:
            estado_atual = STATE_MENU
            status_msg = "Bem-vindo ao Menu Principal."
            status_cor = COR_ACENTO

    # ==========================================
    # ESTADO 3: CÂMERA DE CAPTURA DO COLETOR
    # ==========================================
    elif estado_atual == STATE_CALIB_CAM and cap is not None:
        success, frame_cam = cap.read()
        if not success:
            print(degrade_texto("[!] Erro na captura da câmera.", offset=0.8))
            estado_atual = STATE_CALIB_GRID
            status_msg = "Erro na captura da câmera. Tente novamente."
            status_cor = COR_ERRO
            continue

        detector_async.enviar(frame_cam)
        rid, hands = detector_async.ultimo()
        desenhar_maos(frame_cam, hands)
        hands = ordenar_maos(hands)

        h_cam = cabecalho(canvas, "COLETOR DE SINAIS", f"Modo: {modo_camera_acao}")
        colar_video(canvas, frame_cam, 30, 104, 880, 634, COR_OK if hands else COR_BORDA)

        cartao(canvas, 900, 104, 1250, 634)
        qtd_atual = counts.get(str(label_selecionada), 0)

        cartao(canvas, 920, 124, 1230, 214, COR_CARD_ALT, COR_ACENTO, 12, 2)
        cartao(canvas, 932, 136, 1012, 202, COR_CARD, COR_BORDA, 10)
        txt(972, 169, label_selecionada, 44, COR_TEXTO, True, 'm')
        txt(1028, 150, "DINÂMICO" if tipo_captura_movimento else "ESTÁTICO", 15,
            COR_AVISO if tipo_captura_movimento else COR_TEXTO, True)
        txt(1028, 175, f"{qtd_atual} amostras salvas", 13, COR_TEXTO2)
        txt(1028, 195, "(movimento)" if tipo_captura_movimento else "(sinal parado)", 12, COR_MUDO)

        cv2.circle(canvas, (932, 244), 6, COR_OK if hands else COR_AVISO, -1, cv2.LINE_AA)
        txt(946, 244, "Mão detectada" if hands else "Aguardando mão...", 15, COR_OK if hands else COR_AVISO, True)

        dica = ("Ao gravar, FAÇA o movimento do sinal durante ~1,6 s."
                if tipo_captura_movimento else
                "Ao gravar, mantenha a mão no formato do sinal por ~1 s (pode variar levemente).")
        for i, linha in enumerate(quebrar_linhas(dica, 14, 310)):
            txt(920, 285 + i * 22, linha, 14, COR_TEXTO2)

        txt(920, 356, "Pronto para gravar", 14, COR_MUDO, True)
        barra(canvas, 920, 372, 1230, 386, 0.0, COR_OK)

        hs = botao(canvas, 920, 420, 1230, 490, "GRAVAR SINAL", 'ok', 20, atalho="S")
        hv = botao(canvas, 920, 566, 1230, 614, "VOLTAR AO PAINEL", 'neutro', 16)

        # A tecla usada aqui é a lida no FINAL da iteração ANTERIOR, evitando
        # duas chamadas concorrentes a cv2.waitKey no mesmo frame.
        salvar_agora = (mouse_clicked and hs) or (tecla_atual in (ord('s'), ord('S')))

        if salvar_agora:
            if hands:
                QTD_ALVO = 30
                amostras = []
                cancelado = False

                status_msg = ("Gravando... faça o movimento do sinal!" if tipo_captura_movimento
                              else "Gravando... mantenha a mão no sinal (pode variar levemente).")
                status_cor = COR_OK
                flush_texto(canvas)  # fixa o texto atual antes do sub-loop de gravação
                apresentar(canvas)
                cv2.waitKey(1)

                duracao_captura = 1.6 if tipo_captura_movimento else 1.0
                inicio_captura = time.time()
                while time.time() - inicio_captura < duracao_captura and len(amostras) < QTD_ALVO * 2:
                    ret_sub, frame_sub = cap.read()
                    if not ret_sub:
                        break
                    hands_sub = ordenar_maos(detector_async.detectar_sync(frame_sub))
                    desenhar_maos(frame_sub, hands_sub)

                    if hands_sub:
                        atualizar_buffers_movimento(hands_sub, buffer_movimento_calibracao)
                        mov_lista = features_movimento_das_maos(buffer_movimento_calibracao)
                        amostras.append(extrair_features_avancadas(hands_sub, mov_lista))

                    colar_video(canvas, frame_sub, 30, 104, 880, 634, COR_ERRO)
                    cv2.rectangle(canvas, (920, 344), (1230, 368), COR_CARD, -1)
                    txt(920, 356, "● GRAVANDO...", 14, COR_ERRO, True)
                    barra(canvas, 920, 372, 1230, 386, (time.time() - inicio_captura) / duracao_captura, COR_ERRO)
                    flush_texto(canvas)
                    apresentar(canvas)
                    if (cv2.waitKey(1) & 0xFF) == 27:  # permite cancelar com ESC durante a gravação
                        cancelado = True
                        break

                # Se poucas amostras reais foram capturadas (mão saiu do
                # quadro, gravação muito curta etc.), completa com pequenas
                # variações sintéticas em torno da última amostra válida.
                if amostras and not cancelado and len(amostras) < 10:
                    base = amostras[-1]
                    while len(amostras) < QTD_ALVO:
                        ruido = np.random.normal(0, 0.012, len(base))
                        amostras.append([round(v + r, 4) if v != 0.0 else 0.0 for v, r in zip(base, ruido)])

                if cancelado:
                    status_msg = "Gravação cancelada. Nada foi alterado."
                    status_cor = COR_AVISO
                elif amostras:
                    if modo_camera_acao == 'EDITAR':
                        apagar_amostras_label(label_selecionada)  # só apaga as antigas agora que há novas
                    with open(NOME_ARQUIVO, mode='a', newline='', encoding='utf-8') as f:
                        writer = csv.writer(f)
                        for row in amostras:
                            if len(row) == NUM_FEATURES:
                                writer.writerow([str(label_selecionada)] + row)

                    recarregar_counts = True
                    iniciar_treino_background()
                    status_msg = f"{len(amostras)} amostras salvas para '{label_selecionada}' com sucesso!"
                    status_cor = COR_OK
                    estado_atual = STATE_CALIB_GRID
                else:
                    status_msg = "Não foi possível capturar amostras. Tente novamente."
                    status_cor = COR_ERRO
            else:
                status_msg = "Posicione a mão na tela antes de gravar!"
                status_cor = COR_ERRO

        if mouse_clicked and hv:
            estado_atual = STATE_CALIB_GRID

    # ==========================================
    # ESTADO 4: MODO ESCRITOR / IDENTIFICADOR
    # ==========================================
    elif estado_atual == STATE_WRITER_CAM and cap is not None:
        success, frame_cam = cap.read()
        if not success:
            print(degrade_texto("[!] Falha na imagem.", offset=0.8))
            estado_atual = STATE_MENU
            status_msg = "Falha ao ler a câmera. Voltando ao menu."
            status_cor = COR_ERRO
            continue

        detector_async.enviar(frame_cam)
        rid, hands = detector_async.ultimo()
        desenhar_maos(frame_cam, hands)
        hands = ordenar_maos(hands)
        agora = time.time()

        if hands:
            resultado = None
            if rid != rid_buffer_escritor:
                # trajetória e predição só avançam quando há detecção nova
                # (o vídeo roda mais rápido que o detector)
                rid_buffer_escritor = rid
                atualizar_buffers_movimento(hands, buffer_movimento_escritor)
                mov_lista = features_movimento_das_maos(buffer_movimento_escritor)
                resultado = floresta.prever(extrair_features_avancadas(hands, mov_lista))

            # o buffer de votos só avança quando há predição nova (os demais
            # frames reaproveitam a última letra estável)
            if resultado is not None:
                predicao_instantanea, confianca = resultado

                if confianca > CONFIDENCE_THRESHOLD:
                    buffer_predicoes.append(predicao_instantanea)
                else:
                    buffer_predicoes.append("Identificando...")

                if buffer_predicoes:
                    candidato = max(set(buffer_predicoes), key=buffer_predicoes.count)
                    votos_candidato = buffer_predicoes.count(candidato)
                    # Só aceita como "estável" se houver maioria consistente no
                    # buffer — reduz tremulação entre letras parecidas.
                    predicao_estavel = candidato if votos_candidato >= MIN_VOTOS_ESTAVEL else "Identificando..."
                else:
                    predicao_estavel = "Identificando..."

            if predicao_estavel != "Identificando...":
                if predicao_estavel != ultima_letra_detectada:
                    ultima_letra_detectada = predicao_estavel
                    tempo_inicio_hold = agora
                    letra_adicionada = False
                else:
                    tempo_decorrido = agora - tempo_inicio_hold
                    if tempo_decorrido >= TEMPO_HOLD_NECESSARIO and not letra_adicionada:
                        texto_construido += predicao_estavel
                        letra_adicionada = True
                        tempo_ultima_letra_adicionada = agora
                        espaco_pendente = True
            else:
                ultima_letra_detectada = None
                letra_adicionada = False
        else:
            buffer_predicoes.clear()
            buffer_movimento_escritor[0].clear()
            buffer_movimento_escritor[1].clear()
            ultima_letra_detectada = None
            letra_adicionada = False
            predicao_estavel = "Identificando..."
            confianca = 0.0

        if espaco_pendente and auto_espaco_ativo:
            palavra_atual = texto_construido.split(' ')[-1].strip()
            e_valida = corretor.e_palavra_valida(palavra_atual)
            tempo_limite = TEMPO_AUTO_ESPACO_VALIDO if e_valida else TEMPO_AUTO_ESPACO_INVALIDO

            if (agora - tempo_ultima_letra_adicionada) >= tempo_limite:
                if e_valida and len(texto_construido) > 0 and not texto_construido.endswith(" "):
                    texto_construido += " "
                    status_msg = f"Espaço automático (palavra '{palavra_atual}' confirmada)!"
                    status_cor = COR_OK
                espaco_pendente = False

        # ---------------- UI do Escritor ----------------
        h_cam = cabecalho(canvas, "MODO ESCRITOR", "Identificação em tempo real")

        cartao(canvas, 30, 104, 820, 634)
        txt(54, 128, "TEXTO GERADO", 13, COR_MUDO, True)
        txt(796, 128, f"{len(texto_construido)} caracteres", 13, COR_MUDO, False, 'r')
        cv2.line(canvas, (46, 148), (804, 148), COR_BORDA, 1)
        if texto_construido:
            desenhar_texto_multilinhas(canvas, texto_construido, 54, 184)
        else:
            txt(54, 184, "Faça os sinais na câmera para começar a escrever...", 18, COR_MUDO)
            if int(time.time() * 2) % 2 == 0:
                cv2.line(canvas, (57, 170), (57, 198), COR_ACENTO, 2, cv2.LINE_AA)
        cv2.line(canvas, (46, 600), (804, 600), COR_BORDA, 1)
        txt(54, 618, "[Espaço] espaço   [Backspace] apagar   [C] limpar   [A] auto-espaço   [O] ouvir   [ESC] menu",
            12, COR_MUDO)

        colar_video(canvas, frame_cam, 860, 104, 1240, 318, COR_OK if hands else COR_BORDA)

        cartao(canvas, 860, 334, 1240, 470)
        txt(880, 354, "DETECÇÃO ATUAL", 13, COR_MUDO, True)
        reconhecido = bool(hands) and predicao_estavel != "Identificando..."
        cartao(canvas, 880, 370, 970, 454, COR_CARD_ALT, COR_OK if reconhecido else COR_BORDA, 12, 2 if reconhecido else 1)
        txt(925, 413, predicao_estavel if reconhecido else "—", 48, COR_OK if reconhecido else COR_MUDO, True, 'm')

        txt(990, 382, "Precisão", 13, COR_TEXTO2)
        txt(1220, 382, f"{confianca:.0f}%" if reconhecido else "--", 20, COR_TEXTO if reconhecido else COR_MUDO, True, 'r')
        barra(canvas, 990, 396, 1220, 406, confianca / 100.0 if reconhecido else 0.0, COR_ACENTO2)
        txt(990, 424, "Segure para confirmar", 13, COR_TEXTO2)
        progresso_val = 0.0
        if reconhecido and ultima_letra_detectada:
            progresso_val = 1.0 if letra_adicionada else min(1.0, (agora - tempo_inicio_hold) / TEMPO_HOLD_NECESSARIO)
        barra(canvas, 990, 438, 1220, 448, progresso_val, COR_OK)

        h_ae = botao(canvas, 860, 486, 1240, 530,
                     f"AUTO-ESPAÇO: {'LIGADO' if auto_espaco_ativo else 'DESLIGADO'}",
                     'ok' if auto_espaco_ativo else 'neutro', 15, atalho="A")
        h_ouvir = botao(canvas, 860, 538, 1240, 582, "OUVIR TEXTO", 'azul', 15, atalho="O")
        h_sair = botao(canvas, 860, 590, 1240, 634, "VOLTAR AO MENU", 'neutro', 15, atalho="ESC")

        if mouse_clicked and h_ae:
            auto_espaco_ativo = not auto_espaco_ativo
            status_msg = f"Auto-Espaço {'ativado' if auto_espaco_ativo else 'desativado'}!"
            status_cor = COR_OK if auto_espaco_ativo else COR_AVISO

        if mouse_clicked and h_ouvir:
            falar_texto_async(texto_construido)
            status_msg = "Reproduzindo áudio do texto..."
            status_cor = COR_ACENTO2

        if mouse_clicked and h_sair:
            estado_atual = STATE_MENU

    # Rodapé universal de status + render do texto da tela
    rodape(canvas)
    flush_texto(canvas)

    if mouse_clicked and h_cam:
        modal_camera = True

    if modal_camera:
        cv2.convertScaleAbs(canvas, dst=canvas, alpha=0.35)  # escurece a tela de baixo
        n_cams = len(cameras_disponiveis)
        linha_h, gap_l, larg = 50, 8, 640
        alt = 130 + n_cams * (linha_h + gap_l) + 56
        mx1, my1 = (LARGURA_CANVAS - larg) // 2, max(20, (ALTURA_CANVAS - alt) // 2)
        mx2, my2 = mx1 + larg, my1 + alt
        cartao(canvas, mx1, my1, mx2, my2, COR_CARD, COR_ACENTO, 16, 2)
        txt(mx1 + 28, my1 + 34, "ESCOLHER CÂMERA", 20, COR_TEXTO, True)
        txt(mx1 + 28, my1 + 64, "Clique numa câmera (ou digite o número). ESC ou clique fora para fechar.", 13, COR_TEXTO2)
        y_item = my1 + 90
        for i in cameras_disponiveis:
            em_uso = i == cam_index
            rotulo = f"{i}:  {nome_camera(i)}" + ("    ·  em uso" if em_uso else "")
            if botao(canvas, mx1 + 24, y_item, mx2 - 24, y_item + linha_h, rotulo,
                     'ok' if em_uso else 'neutro', 16, atalho=str(i)) and clique_modal:
                camera_escolhida = i
            y_item += linha_h + gap_l
        h_atualizar = FilterGraph is not None and botao(
            canvas, mx1 + 24, my2 - 56, mx1 + 24 + 230, my2 - 16, "ATUALIZAR LISTA", 'azul', 14)
        h_fechar = botao(canvas, mx2 - 24 - 150, my2 - 56, mx2 - 24, my2 - 16, "FECHAR", 'neutro', 14, atalho="ESC")
        dentro = mx1 <= mouse_pos[0] <= mx2 and my1 <= mouse_pos[1] <= my2
        if clique_modal and camera_escolhida is None:
            if h_atualizar:
                atualizar_lista_cameras()
                status_msg = f"Lista atualizada: {len(cameras_disponiveis)} câmera(s) encontrada(s)."
                status_cor = COR_OK
            elif h_fechar or not dentro:
                modal_camera = False
        flush_texto(canvas)

    if camera_escolhida is not None:
        escolhida, camera_escolhida, modal_camera = camera_escolhida, None, False
        if escolhida == cam_index:
            status_msg = f"A câmera {cam_index}: {nome_camera(cam_index)} já está em uso."
            status_cor = COR_ACENTO
        else:
            cartao(canvas, 440, 300, 840, 400, COR_CARD, COR_ACENTO, 14, 2)
            txt(640, 335, "Trocando de câmera...", 22, COR_TEXTO, True, 'm')
            txt(640, 368, f"Abrindo {nome_camera(escolhida)}", 14, COR_TEXTO2, False, 'm')
            flush_texto(canvas)
            apresentar(canvas)
            cv2.waitKey(1)
            if trocar_camera(escolhida):
                status_msg = f"Câmera {cam_index}: {nome_camera(cam_index)} em uso."
                status_cor = COR_OK
            else:
                status_msg = f"Não abriu '{nome_camera(escolhida)}' (DroidCam: conecte o cliente ao celular). Mantida a atual."
                status_cor = COR_ERRO

    mouse_clicked = False

    apresentar(canvas)
    key = cv2.waitKey(1) & 0xFF

    # Janela fechada pelo "X"
    if cv2.getWindowProperty(NOME_JANELA, cv2.WND_PROP_VISIBLE) < 1:
        break

    if modal_camera and ord('0') <= key <= ord('9') and (key - ord('0')) in cameras_disponiveis:
        camera_escolhida = key - ord('0')

    if estado_atual == STATE_WRITER_CAM and not modal_camera:
        if key == 8:  # Backspace
            texto_construido = texto_construido[:-1]
            espaco_pendente = False
        elif key in (ord('c'), ord('C')):  # Limpar texto
            texto_construido = ""
            espaco_pendente = False
        elif key == 32:  # Espaço manual
            texto_construido += " "
            espaco_pendente = False
        elif key in (ord('a'), ord('A')):  # Alternar Auto-Espaço
            auto_espaco_ativo = not auto_espaco_ativo
            status_msg = f"Auto-Espaço {'ativado' if auto_espaco_ativo else 'desativado'}!"
            status_cor = COR_OK if auto_espaco_ativo else COR_AVISO
        elif key in (ord('o'), ord('O')):  # Ouvir Texto (TTS)
            falar_texto_async(texto_construido)
            status_msg = "Reproduzindo áudio do texto..."
            status_cor = COR_ACENTO2

    if key == 9:  # Tab: abre/fecha o seletor de câmera
        modal_camera = not modal_camera

    if key == 27:
        if modal_camera:
            modal_camera = False
        elif estado_atual != STATE_MENU:
            # ESC na ação volta ao grid; no grid/câmera/escritor volta ao menu
            estado_atual = STATE_CALIB_GRID if estado_atual in (STATE_CALIB_ACTION, STATE_CALIB_CAM) else STATE_MENU
        else:
            break

    tecla_atual = key  # disponível para a PRÓXIMA iteração (ex.: atalho [S])

if cap is not None:
    cap.release()
cv2.destroyAllWindows()

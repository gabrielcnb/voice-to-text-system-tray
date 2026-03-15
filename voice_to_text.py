"""
Voice to Text - System Tray

Roda em background no system tray.
Segure CTRL DIREITO para gravar, solte para transcrever.
"""

# ============== SISTEMA DE LOGS ==============
import subprocess
import sys
import os
import logging
from datetime import datetime

# Configurar pasta de logs
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(SCRIPT_DIR, "voice_to_text.log")

# Configurar logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger("VoiceToText")

log.info("=" * 50)
log.info("VOICE TO TEXT - INICIANDO")
log.info("=" * 50)
log.info(f"Python: {sys.version}")
log.info(f"Executavel: {sys.executable}")
log.info(f"Diretorio: {SCRIPT_DIR}")

# ============== VERIFICAR INSTANCIA UNICA ==============
import ctypes

PID_FILE = os.path.join(SCRIPT_DIR, ".voicetotext.pid")

def _fechar_instancia_anterior():
    """Fecha a instância anterior se existir"""
    if os.path.exists(PID_FILE):
        try:
            with open(PID_FILE, "r") as f:
                old_pid = int(f.read().strip())

            # Tentar matar o processo anterior
            import signal
            os.kill(old_pid, signal.SIGTERM)
            log.info(f"Instancia anterior (PID {old_pid}) encerrada")

            # Aguardar um pouco para o processo fechar
            import time
            time.sleep(0.5)
        except (ValueError, ProcessLookupError, PermissionError, OSError) as e:
            log.debug(f"Processo anterior ja encerrado ou inacessivel: {e}")
        except Exception as e:
            log.debug(f"Erro ao fechar instancia anterior: {e}")

    # Salvar nosso PID
    with open(PID_FILE, "w") as f:
        f.write(str(os.getpid()))
    log.info(f"PID atual: {os.getpid()}")

_fechar_instancia_anterior()
log.info("Instancia unica verificada")
# ========================================================


def _verificar_admin():
    """Verifica se esta rodando como administrador"""
    try:
        import ctypes
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
        log.info(f"Rodando como administrador: {is_admin}")
        return is_admin
    except:
        log.warning("Nao foi possivel verificar permissoes de admin")
        return False


def _mostrar_erro_admin():
    """Mostra mensagem sobre necessidade de admin"""
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()

    messagebox.showwarning(
        "Voice to Text - Aviso",
        "Para funcionar corretamente, execute como Administrador.\n\n"
        "Clique com botao direito no arquivo e selecione:\n"
        "'Executar como administrador'\n\n"
        "Isso e necessario para capturar teclas globalmente."
    )
    root.destroy()


# ============== AUTO-INSTALADOR ==============

# URL do Python 3.12 (versao estavel e compativel)
PYTHON_312_URL = "https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe"
PYTHON_312_INSTALLER = "python-3.12.8-amd64.exe"


def _baixar_python_312(janela_progresso=None, label_status=None, progress_bar=None):
    """Baixa o instalador do Python 3.12"""
    import urllib.request

    destino = os.path.join(SCRIPT_DIR, PYTHON_312_INSTALLER)

    log.info(f"Baixando Python 3.12 de {PYTHON_312_URL}...")

    try:
        def progresso(count, block_size, total_size):
            if progress_bar and total_size > 0:
                percent = int(count * block_size * 100 / total_size)
                progress_bar['value'] = min(percent, 100)
                if label_status:
                    mb_baixado = (count * block_size) / (1024 * 1024)
                    mb_total = total_size / (1024 * 1024)
                    label_status.config(text=f"Baixando... {mb_baixado:.1f} / {mb_total:.1f} MB")
                if janela_progresso:
                    janela_progresso.update()

        urllib.request.urlretrieve(PYTHON_312_URL, destino, progresso)
        log.info(f"Download concluido: {destino}")
        return destino

    except Exception as e:
        log.error(f"Erro ao baixar Python: {e}")
        return None


def _instalar_python_312(instalador_path):
    """Instala o Python 3.12 silenciosamente"""
    log.info("Instalando Python 3.12...")

    # Parametros para instalacao silenciosa
    # /quiet = silencioso
    # /passive = mostra progresso mas sem interacao
    # PrependPath=1 = adiciona ao PATH
    # InstallAllUsers=0 = instala so para o usuario atual
    cmd = [
        instalador_path,
        "/passive",
        "InstallAllUsers=0",
        "PrependPath=1",
        "Include_test=0",
        "Include_doc=0"
    ]

    log.info(f"Executando: {' '.join(cmd)}")

    try:
        result = subprocess.run(cmd, capture_output=True, text=True)
        log.info(f"Instalacao finalizada com codigo: {result.returncode}")
        return result.returncode == 0
    except Exception as e:
        log.error(f"Erro ao instalar: {e}")
        return False


def _encontrar_python_312():
    """Procura o Python 3.12 instalado"""
    possiveis_caminhos = [
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Python\Python312\python.exe"),
        r"C:\Python312\pythonw.exe",
        r"C:\Python312\python.exe",
        os.path.expandvars(r"%PROGRAMFILES%\Python312\pythonw.exe"),
        os.path.expandvars(r"%PROGRAMFILES%\Python312\python.exe"),
    ]

    for caminho in possiveis_caminhos:
        if os.path.exists(caminho):
            log.info(f"Python 3.12 encontrado: {caminho}")
            return caminho

    return None


def _reiniciar_com_python_312():
    """Reinicia o script usando Python 3.12"""
    python312 = _encontrar_python_312()

    if python312:
        log.info(f"Reiniciando com Python 3.12: {python312}")
        script_path = os.path.abspath(__file__)

        # Usar pythonw se disponivel (sem janela de console)
        if "python.exe" in python312:
            python312 = python312.replace("python.exe", "pythonw.exe")
            if not os.path.exists(python312):
                python312 = python312.replace("pythonw.exe", "python.exe")

        subprocess.Popen([python312, script_path], cwd=SCRIPT_DIR)
        sys.exit(0)
    else:
        log.error("Python 3.12 nao encontrado apos instalacao!")


def _verificar_versao_python():
    """Verifica se a versao do Python e compativel e instala 3.12 se necessario"""
    import tkinter as tk
    from tkinter import ttk, messagebox

    version = sys.version_info
    log.info(f"Versao Python: {version.major}.{version.minor}.{version.micro}")

    # Python 3.8 a 3.12 = OK
    if version.major == 3 and 8 <= version.minor <= 12:
        log.info("Versao do Python compativel!")
        return True

    # Python muito antigo
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        log.error("Python muito antigo!")
        messagebox.showerror(
            "Voice to Text - Erro",
            "Seu Python e muito antigo!\n\n"
            "O programa vai baixar e instalar Python 3.12 automaticamente."
        )

    # Python 3.13+ (muito novo, incompativel)
    if version.major == 3 and version.minor >= 13:
        log.warning(f"Python {version.minor} detectado - versao incompativel!")

        # Verificar se Python 3.12 ja esta instalado
        python312 = _encontrar_python_312()
        if python312:
            log.info("Python 3.12 ja esta instalado! Reiniciando...")
            _reiniciar_com_python_312()
            return False

        # Perguntar se quer instalar
        root = tk.Tk()
        root.withdraw()

        resposta = messagebox.askyesno(
            "Voice to Text - Python Incompativel",
            f"Seu Python {version.major}.{version.minor} e muito novo e incompativel.\n\n"
            "Deseja baixar e instalar Python 3.12 automaticamente?\n\n"
            "Isso levara alguns minutos.\n"
            "Seu Python atual NAO sera desinstalado.",
            icon='warning'
        )
        root.destroy()

        if not resposta:
            log.info("Usuario recusou instalar Python 3.12")
            messagebox.showinfo(
                "Voice to Text",
                "Voce pode instalar Python 3.12 manualmente:\n\n"
                "python.org/downloads/release/python-3128/"
            )
            sys.exit(0)

        # Criar janela de progresso
        root = tk.Tk()
        root.title("Voice to Text - Instalando Python 3.12")
        root.geometry("450x180")
        root.resizable(False, False)
        root.attributes('-topmost', True)

        # Centralizar
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - 225
        y = (root.winfo_screenheight() // 2) - 90
        root.geometry(f"+{x}+{y}")

        tk.Label(root, text="Baixando Python 3.12...",
                 font=("Segoe UI", 12, "bold")).pack(pady=15)

        progress = ttk.Progressbar(root, length=400, mode='determinate')
        progress.pack(pady=10)

        status = tk.Label(root, text="Iniciando download...", font=("Segoe UI", 9))
        status.pack()

        info = tk.Label(root, text="Isso pode levar alguns minutos",
                        font=("Segoe UI", 8), fg="#888")
        info.pack(pady=5)

        root.update()

        # Baixar Python
        instalador = _baixar_python_312(root, status, progress)

        if not instalador:
            root.destroy()
            messagebox.showerror(
                "Voice to Text - Erro",
                "Falha ao baixar Python 3.12.\n\n"
                "Verifique sua conexao com a internet."
            )
            sys.exit(1)

        # Instalar Python
        status.config(text="Instalando Python 3.12... (aguarde)")
        progress['mode'] = 'indeterminate'
        progress.start(10)
        root.update()

        sucesso = _instalar_python_312(instalador)

        progress.stop()
        root.destroy()

        # Limpar instalador
        try:
            os.remove(instalador)
            log.info("Instalador removido")
        except:
            pass

        if sucesso:
            log.info("Python 3.12 instalado com sucesso!")
            messagebox.showinfo(
                "Voice to Text",
                "Python 3.12 instalado com sucesso!\n\n"
                "O programa vai reiniciar automaticamente."
            )
            _reiniciar_com_python_312()
        else:
            messagebox.showerror(
                "Voice to Text - Erro",
                "Falha ao instalar Python 3.12.\n\n"
                "Tente instalar manualmente:\n"
                "python.org/downloads/release/python-3128/"
            )
            sys.exit(1)

        return False

    return True


def _verificar_dependencias():
    """Verifica e instala dependencias automaticamente na primeira execucao"""
    log.info("Verificando dependencias...")

    # Verificar versao do Python primeiro (instala 3.12 se necessario)
    if not _verificar_versao_python():
        sys.exit(0)  # Vai reiniciar com Python correto

    MARKER = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".installed")

    if os.path.exists(MARKER):
        log.info("Dependencias ja instaladas (marker encontrado)")
        return

    DEPS = [
        ("speech_recognition", "SpeechRecognition"),
        ("keyboard", "keyboard"),
        ("pyautogui", "pyautogui"),
        ("pyperclip", "pyperclip"),
        ("pystray", "pystray"),
        ("PIL", "pillow"),
    ]

    faltando = []
    for modulo, pip_pkg in DEPS:
        try:
            __import__(modulo)
            log.debug(f"  [OK] {modulo}")
        except ImportError:
            log.warning(f"  [FALTA] {modulo}")
            faltando.append(pip_pkg)

    # Verificar PyAudio separadamente
    try:
        import pyaudio
        log.debug("  [OK] pyaudio")
    except ImportError:
        log.warning("  [FALTA] pyaudio")
        faltando.append("PyAudio")

    if not faltando:
        log.info("Todas dependencias OK")
        open(MARKER, 'w').close()
        return

    log.info(f"Instalando {len(faltando)} dependencias faltantes...")
    _instalar_com_gui(faltando, MARKER)

def _instalar_com_gui(pacotes, marker):
    """Instala pacotes mostrando janela de progresso"""
    import tkinter as tk
    from tkinter import ttk

    log.info("Abrindo janela de instalacao...")

    root = tk.Tk()
    root.title("Voice to Text - Instalando...")
    root.geometry("420x160")
    root.resizable(False, False)
    root.attributes('-topmost', True)

    # Centralizar janela
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - 210
    y = (root.winfo_screenheight() // 2) - 80
    root.geometry(f"+{x}+{y}")

    tk.Label(root, text="Primeira execucao - instalando dependencias...",
             font=("Segoe UI", 11)).pack(pady=15)

    progress = ttk.Progressbar(root, length=380, mode='determinate')
    progress.pack(pady=10)

    status = tk.Label(root, text="", font=("Segoe UI", 9), fg="#666")
    status.pack()

    def instalar():
        total = len(pacotes)
        pyaudio_falhou = False

        for i, pkg in enumerate(pacotes):
            nome = pkg.split('==')[0]
            log.info(f"Instalando: {pkg}")
            status.config(text=f"Instalando {nome}...")
            progress['value'] = ((i + 1) / (total + 1)) * 100
            root.update()

            if pkg == "PyAudio":
                # PyAudio precisa tratamento especial no Windows
                log.info("Tentando instalar PyAudio...")

                # Metodo 1: pip install normal
                result = subprocess.run(
                    [sys.executable, "-m", "pip", "install", "PyAudio"],
                    capture_output=True, text=True
                )
                if result.returncode == 0:
                    log.info("PyAudio instalado com sucesso!")
                    continue

                log.warning(f"PyAudio pip falhou")

                # Metodo 2: Tentar PyAudio-wheels (tem versoes pre-compiladas)
                log.info("Tentando PyAudio-wheels...")
                result = subprocess.run(
                    [sys.executable, "-m", "pip", "install", "PyAudio-wheels"],
                    capture_output=True, text=True
                )
                if result.returncode == 0:
                    log.info("PyAudio-wheels instalado com sucesso!")
                    continue

                # Metodo 3: Tentar versao especifica do PyPI
                log.info("Tentando versao especifica...")
                result = subprocess.run(
                    [sys.executable, "-m", "pip", "install", "--only-binary=:all:", "PyAudio"],
                    capture_output=True, text=True
                )
                if result.returncode == 0:
                    log.info("PyAudio binario instalado!")
                    continue

                log.error("PyAudio nao pode ser instalado automaticamente")
                log.error("Python 3.13+ pode ter problemas de compatibilidade")
                pyaudio_falhou = True

            else:
                result = subprocess.run(
                    [sys.executable, "-m", "pip", "install", pkg],
                    capture_output=True, text=True
                )
                if result.returncode == 0:
                    log.info(f"  {nome} instalado com sucesso")
                else:
                    log.error(f"  {nome} falhou: {result.stderr}")

        progress['value'] = 100

        if pyaudio_falhou:
            status.config(text="Aviso: PyAudio nao instalado", fg="orange")
            log.warning("=" * 50)
            log.warning("PyAudio nao pode ser instalado!")
            log.warning("Seu Python pode ser muito novo (3.13+)")
            log.warning("Recomendamos instalar Python 3.11 ou 3.12")
            log.warning("Download: python.org/downloads")
            log.warning("=" * 50)
            root.update()

            from tkinter import messagebox
            messagebox.showwarning(
                "Voice to Text - Aviso",
                "PyAudio nao pode ser instalado automaticamente.\n\n"
                "Seu Python (3.13+) e muito novo e pode ter\n"
                "problemas de compatibilidade.\n\n"
                "SOLUCAO: Instale Python 3.11 ou 3.12\n"
                "Download: python.org/downloads\n\n"
                "O programa tentara continuar mesmo assim..."
            )
        else:
            status.config(text="Concluido! Iniciando aplicativo...")
            root.update()

        open(marker, 'w').close()
        log.info("Instalacao concluida!")
        root.after(800, root.destroy)

    root.after(100, instalar)
    root.mainloop()

_verificar_dependencias()
# =============================================

log.info("Importando bibliotecas...")

try:
    import speech_recognition as sr
    log.debug("  speech_recognition OK")
except Exception as e:
    log.error(f"  speech_recognition ERRO: {e}")
    raise

try:
    import keyboard
    log.debug("  keyboard OK")
except Exception as e:
    log.error(f"  keyboard ERRO: {e}")
    log.error("  DICA: Execute como Administrador!")
    raise

import threading
import time
import ctypes
import wave
import struct
import math
import tempfile

try:
    import pyperclip
    log.debug("  pyperclip OK")
except Exception as e:
    log.error(f"  pyperclip ERRO: {e}")
    raise

try:
    import pyautogui
    log.debug("  pyautogui OK")
except Exception as e:
    log.error(f"  pyautogui ERRO: {e}")
    raise

try:
    import pystray
    log.debug("  pystray OK")
except Exception as e:
    log.error(f"  pystray ERRO: {e}")
    raise

from PIL import Image, ImageDraw
log.info("Todas bibliotecas importadas com sucesso!")

# ============== CONFIGURACOES ==============
import json

CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")
MAX_HISTORICO = 10

# Teclas disponiveis (nome_interno, nome_exibicao)
TECLAS_DISPONIVEIS = [
    ("right ctrl", "CTRL Direito"),
    ("left ctrl", "CTRL Esquerdo"),
    ("right shift", "SHIFT Direito"),
    ("right alt", "ALT Direito"),
    ("f8", "F8"),
    ("f9", "F9"),
    ("f10", "F10"),
    ("scroll lock", "Scroll Lock"),
    ("pause", "Pause"),
]

# Configuracoes padrao
CONFIG_PADRAO = {
    "tecla": "right ctrl",
    "tecla_display": "CTRL Direito",
    "colar_automatico": True
}


def carregar_config():
    """Carrega configuracoes do arquivo"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                log.info(f"Configuracoes carregadas: {config}")
                return config
    except Exception as e:
        log.error(f"Erro ao carregar config: {e}")
    return None


def salvar_config(config):
    """Salva configuracoes no arquivo"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        log.info(f"Configuracoes salvas: {config}")
    except Exception as e:
        log.error(f"Erro ao salvar config: {e}")


def mostrar_config_inicial():
    """Mostra janela para configuracao inicial"""
    import tkinter as tk
    from tkinter import ttk

    config = CONFIG_PADRAO.copy()

    root = tk.Tk()
    root.title("Voice to Text - Configuracao")
    root.geometry("400x300")
    root.resizable(False, False)
    root.attributes('-topmost', True)

    # Centralizar
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - 200
    y = (root.winfo_screenheight() // 2) - 150
    root.geometry(f"+{x}+{y}")

    tk.Label(root, text="Configuracao Inicial",
             font=("Segoe UI", 14, "bold")).pack(pady=15)

    # Frame para tecla
    frame_tecla = tk.Frame(root)
    frame_tecla.pack(pady=10, padx=20, fill="x")

    tk.Label(frame_tecla, text="Tecla para gravar:",
             font=("Segoe UI", 10)).pack(anchor="w")

    tecla_var = tk.StringVar(value="CTRL Direito")
    combo_tecla = ttk.Combobox(frame_tecla, textvariable=tecla_var,
                                values=[t[1] for t in TECLAS_DISPONIVEIS],
                                state="readonly", font=("Segoe UI", 10))
    combo_tecla.pack(fill="x", pady=5)

    # Frame para colar automatico
    frame_colar = tk.Frame(root)
    frame_colar.pack(pady=10, padx=20, fill="x")

    colar_var = tk.BooleanVar(value=True)
    tk.Checkbutton(frame_colar, text="Colar texto automaticamente",
                   variable=colar_var, font=("Segoe UI", 10)).pack(anchor="w")

    tk.Label(frame_colar, text="(Se desativado, apenas copia para a area de transferencia)",
             font=("Segoe UI", 8), fg="#666").pack(anchor="w")

    # Instrucao
    tk.Label(root, text="\nSegure a tecla para gravar, solte para transcrever.",
             font=("Segoe UI", 9), fg="#444").pack()

    def confirmar():
        # Encontrar tecla selecionada
        nome_exibicao = tecla_var.get()
        for interno, exibicao in TECLAS_DISPONIVEIS:
            if exibicao == nome_exibicao:
                config["tecla"] = interno
                config["tecla_display"] = exibicao
                break

        config["colar_automatico"] = colar_var.get()
        salvar_config(config)
        root.destroy()

    tk.Button(root, text="Confirmar", command=confirmar,
              font=("Segoe UI", 11), width=15, bg="#4CAF50", fg="white").pack(pady=20)

    root.protocol("WM_DELETE_WINDOW", confirmar)  # Salva ao fechar
    root.mainloop()

    return config


# Carregar ou criar configuracao
CONFIG = carregar_config()
if CONFIG is None:
    log.info("Primeira execucao - mostrando configuracao inicial")
    CONFIG = mostrar_config_inicial()

TECLA = CONFIG.get("tecla", "right ctrl")
TECLA_DISPLAY = CONFIG.get("tecla_display", "CTRL Direito")
COLAR_AUTOMATICO = CONFIG.get("colar_automatico", True)

log.info(f"Tecla: {TECLA_DISPLAY} ({TECLA})")
log.info(f"Colar automatico: {COLAR_AUTOMATICO}")
# ===========================================

user32 = ctypes.windll.user32

# Pasta temporaria para sons
TEMP_DIR = tempfile.gettempdir()


def gerar_som_wav(nome, frequencias, duracao=0.15, volume=0.3):
    """Gera um arquivo WAV com tons agradaveis"""
    filepath = os.path.join(TEMP_DIR, f"vtt_{nome}.wav")

    if os.path.exists(filepath):
        return filepath

    sample_rate = 44100
    n_samples = int(sample_rate * duracao)

    # Gerar som com multiplas frequencias (acorde)
    samples = []
    for i in range(n_samples):
        t = i / sample_rate
        # Envelope (fade in/out suave)
        envelope = math.sin(math.pi * i / n_samples)

        # Mixar frequencias
        valor = 0
        for freq in frequencias:
            valor += math.sin(2 * math.pi * freq * t)
        valor = valor / len(frequencias) * envelope * volume

        samples.append(int(valor * 32767))

    # Escrever WAV
    with wave.open(filepath, 'w') as wav:
        wav.setnchannels(1)
        wav.setsampwidth(2)
        wav.setframerate(sample_rate)
        wav.writeframes(struct.pack(f'{len(samples)}h', *samples))

    return filepath


def inicializar_sons():
    """Cria os arquivos de som"""
    global SOM_START, SOM_STOP, SOM_SUCCESS, SOM_ERROR

    # Som de inicio: acorde suave ascendente (C-E)
    SOM_START = gerar_som_wav("start", [523, 659], duracao=0.12, volume=0.25)

    # Som de parada: nota descendente (G)
    SOM_STOP = gerar_som_wav("stop", [392], duracao=0.1, volume=0.2)

    # Som de sucesso: acorde maior alegre (C-E-G)
    SOM_SUCCESS = gerar_som_wav("success", [523, 659, 784], duracao=0.2, volume=0.25)

    # Som de erro: tom grave (A baixo)
    SOM_ERROR = gerar_som_wav("error", [220, 185], duracao=0.25, volume=0.2)


def play_sound(tipo):
    """Toca som customizado"""
    import winsound
    sons = {
        'start': SOM_START,
        'stop': SOM_STOP,
        'success': SOM_SUCCESS,
        'error': SOM_ERROR
    }
    try:
        filepath = sons.get(tipo)
        if filepath and os.path.exists(filepath):
            winsound.PlaySound(filepath, winsound.SND_FILENAME | winsound.SND_ASYNC)
    except:
        pass


def corrigir_gagueira(texto):
    if not texto:
        return texto

    palavras = texto.split()
    resultado = []

    i = 0
    while i < len(palavras):
        palavra_atual = palavras[i].lower().strip('.,!?')
        j = i + 1
        while j < len(palavras) and palavras[j].lower().strip('.,!?') == palavra_atual:
            j += 1
        resultado.append(palavras[i])
        i = j

    return ' '.join(resultado)


def adicionar_pontuacao(texto):
    """Adiciona pontuacao automatica e primeira letra maiuscula"""
    if not texto:
        return texto

    texto = texto.strip()

    # Primeira letra maiuscula
    if len(texto) > 1:
        texto = texto[0].upper() + texto[1:]
    else:
        texto = texto.upper()

    # Se ja tem pontuacao no final, retorna
    if texto[-1] in '.!?':
        return texto

    texto_lower = texto.lower()
    palavras = texto_lower.split()

    # Palavras interrogativas no inicio
    interrogativas_inicio = [
        'o que', 'oque', 'como', 'quando', 'onde', 'por que', 'porque',
        'qual', 'quais', 'quem', 'quanto', 'quantos', 'quantas',
        'será', 'sera', 'seria', 'pode', 'posso', 'podemos',
        'você', 'voce', 'vocês', 'voces', 'tu', 'ele', 'ela',
        'isso', 'aquilo', 'esse', 'essa', 'é', 'e possivel'
    ]

    # Verifica inicio
    for interrog in interrogativas_inicio:
        if texto_lower.startswith(interrog):
            return texto + '?'

    # Palavras interrogativas no final
    interrogativas_fim = [
        'né', 'ne', 'certo', 'sim', 'não', 'nao', 'hein', 'ein',
        'mesmo', 'sabe', 'entende', 'entendeu', 'tá', 'ta', 'ok'
    ]

    if palavras:
        ultima = palavras[-1].strip('.,!?')
        if ultima in interrogativas_fim:
            return texto + '?'

    # Palavras exclamativas
    exclamativas = [
        'nossa', 'caramba', 'uau', 'wow', 'legal', 'incrivel', 'incrível',
        'demais', 'ótimo', 'otimo', 'perfeito', 'excelente', 'maravilhoso',
        'que legal', 'muito bom', 'adorei', 'amei', 'parabéns', 'parabens',
        'obrigado', 'obrigada', 'valeu', 'thanks', 'thank you'
    ]

    for excl in exclamativas:
        if excl in texto_lower:
            return texto + '!'

    # Padrao: ponto final
    return texto + '.'


def criar_icone(cor="green"):
    size = 64
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    cores = {
        "green": "#22c55e",
        "red": "#ef4444",
        "yellow": "#eab308"
    }

    draw.ellipse([4, 4, size-4, size-4], fill=cores.get(cor, "#22c55e"))
    draw.ellipse([22, 16, 42, 40], fill="white")
    draw.rectangle([28, 40, 36, 52], fill="white")

    return img


def resumir_texto(texto, max_len=35):
    if not texto:
        return "(vazio)"
    if len(texto) <= max_len:
        return texto
    return texto[:max_len] + "..."


# ============== OVERLAY DE GRAVACAO ==============

# Nivel de audio global (atualizado pela gravacao)
_audio_level = 0.0

def set_audio_level(level):
    global _audio_level
    _audio_level = level

def get_audio_level():
    global _audio_level
    return _audio_level


class RecordingOverlay:
    """Overlay minimalista que aparece durante gravacao"""

    def __init__(self):
        self.root = None
        self.canvas = None
        self._thread = None
        self._should_show = False
        self._running = True
        self._phase = 0
        self._levels = [0.0] * 20  # Historico de niveis

    def _criar_janela(self):
        """Cria janela do overlay"""
        import tkinter as tk

        self.root = tk.Tk()
        self.root.withdraw()
        self.root.overrideredirect(True)
        self.root.attributes('-topmost', True)
        self.root.attributes('-alpha', 0.9)

        try:
            self.root.attributes('-transparentcolor', '#010101')
        except:
            pass

        # Tamanho compacto
        self.width = 160
        self.height = 50

        # Posicao: centro inferior da tela
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = (screen_w - self.width) // 2
        y = screen_h - 120

        self.root.geometry(f"{self.width}x{self.height}+{x}+{y}")

        self.canvas = tk.Canvas(
            self.root,
            width=self.width,
            height=self.height,
            bg='#010101',
            highlightthickness=0
        )
        self.canvas.pack()

    def _desenhar(self):
        """Desenha o overlay"""
        if not self.canvas:
            return

        self.canvas.delete("all")

        # Fundo pill preto
        r = self.height // 2 - 4
        self.canvas.create_oval(4, 4, 4 + r*2, self.height - 4, fill="#000000", outline="#333333")
        self.canvas.create_oval(self.width - 4 - r*2, 4, self.width - 4, self.height - 4, fill="#000000", outline="#333333")
        self.canvas.create_rectangle(4 + r, 4, self.width - 4 - r, self.height - 4, fill="#000000", outline="")
        self.canvas.create_line(4 + r, 4, self.width - 4 - r, 4, fill="#333333")
        self.canvas.create_line(4 + r, self.height - 4, self.width - 4 - r, self.height - 4, fill="#333333")

        # Ponto vermelho pulsante
        pulse = 0.7 + 0.3 * math.sin(self._phase)
        dot_size = int(8 * pulse)
        cx, cy = 28, self.height // 2
        self.canvas.create_oval(
            cx - dot_size, cy - dot_size,
            cx + dot_size, cy + dot_size,
            fill="#ff3b30", outline=""
        )

        # Barras de audio REAL
        bar_x = 50
        num_bars = len(self._levels)
        bar_spacing = 5
        for i, level in enumerate(self._levels):
            h = max(2, int(level * 16))  # Altura baseada no nivel real
            x = bar_x + i * bar_spacing
            self.canvas.create_line(
                x, cy - h, x, cy + h,
                fill="#ffffff", width=2, capstyle="round"
            )

    def _loop(self):
        """Loop de animacao"""
        if not self._running:
            if self.root:
                self.root.quit()
            return

        if self._should_show:
            if not self.root.winfo_viewable():
                self.root.deiconify()
                self.root.lift()
                self._levels = [0.0] * 20  # Reset ao mostrar

            self._phase += 0.2

            # Atualizar historico com nivel real
            level = get_audio_level()
            self._levels.pop(0)
            self._levels.append(level)

            self._desenhar()
        else:
            if self.root.winfo_viewable():
                self.root.withdraw()

        self.root.after(33, self._loop)

    def _run(self):
        """Thread do overlay"""
        try:
            self._criar_janela()
            self._desenhar()
            self.root.after(33, self._loop)
            self.root.mainloop()
        except Exception as e:
            log.debug(f"Overlay erro: {e}")

    def iniciar(self):
        """Inicia o overlay"""
        if self._thread is None or not self._thread.is_alive():
            self._running = True
            self._thread = threading.Thread(target=self._run, daemon=True)
            self._thread.start()
            time.sleep(0.1)  # Aguardar inicializacao

    def mostrar(self):
        self._should_show = True

    def esconder(self):
        self._should_show = False

    def parar(self):
        self._running = False
        self._should_show = False


_overlay = None

def get_overlay():
    global _overlay
    if _overlay is None:
        _overlay = RecordingOverlay()
        _overlay.iniciar()
    return _overlay


class VoiceToText:
    def __init__(self):
        log.info("Inicializando VoiceToText...")

        self.is_recording = False
        self.is_processing = False
        self.target_window = None
        self.audio_data = None
        self.icon = None
        self.running = True
        self.historico = []
        self.record_start_time = None
        self.colar_automatico = COLAR_AUTOMATICO

        log.debug("Configurando reconhecedor de voz...")
        self.recognizer = sr.Recognizer()
        self.recognizer.energy_threshold = 400
        self.recognizer.dynamic_energy_threshold = False
        self.recognizer.pause_threshold = 0.8

        log.info("Testando microfone...")
        try:
            with sr.Microphone() as source:
                log.debug(f"  Microfone encontrado: {source.SAMPLE_RATE}Hz")
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                log.info("Microfone OK!")
        except OSError as e:
            log.error(f"ERRO ao acessar microfone: {e}")
            log.error("Verifique se o microfone esta conectado e permitido")
            play_sound('error')
            raise
        except Exception as e:
            log.error(f"ERRO inesperado no microfone: {e}")
            play_sound('error')
            raise

    def get_active_window(self):
        return user32.GetForegroundWindow()

    def set_active_window(self, hwnd):
        if hwnd:
            try:
                user32.SetForegroundWindow(hwnd)
                time.sleep(0.1)
            except:
                pass

    def paste_text(self, text):
        if not text:
            return False

        # Sempre copia para clipboard
        pyperclip.copy(text)

        # Se colar automatico esta desativado, apenas copia
        if not self.colar_automatico:
            log.info("Texto copiado para clipboard (colar automatico desativado)")
            return True

        # Colar automaticamente
        self.set_active_window(self.target_window)
        time.sleep(0.15)

        pyautogui.hotkey('ctrl', 'v')

        return True

    def copiar_do_historico(self, texto):
        def action(icon, item):
            pyperclip.copy(texto)
            play_sound('success')
        return action

    def atualizar_menu(self):
        items = [
            pystray.MenuItem(f"Tecla: {TECLA_DISPLAY}", lambda: None, enabled=False),
            pystray.MenuItem("─────────────", lambda: None, enabled=False),
        ]

        if self.historico:
            items.append(pystray.MenuItem("Historico (clique p/ copiar):", lambda: None, enabled=False))
            for texto in self.historico[-5:]:
                resumo = resumir_texto(texto)
                items.append(pystray.MenuItem(f"  {resumo}", self.copiar_do_historico(texto)))
            items.append(pystray.MenuItem("─────────────", lambda: None, enabled=False))

        items.append(pystray.MenuItem("Sair", self.quit_app))

        self.icon.menu = pystray.Menu(*items)

    def update_icon(self, status):
        if self.icon:
            cores = {"ready": "green", "recording": "red", "processing": "yellow"}
            tooltips = {
                "ready": "Segure CTRL DIREITO para gravar",
                "recording": "GRAVANDO... solte para transcrever",
                "processing": "Processando audio..."
            }
            self.icon.icon = criar_icone(cores.get(status, "green"))
            self.icon.title = tooltips.get(status, "Voice to Text")

    def do_record(self):
        log.debug("Iniciando gravacao...")
        self.update_icon("recording")

        try:
            import pyaudio

            RATE = 16000
            CHUNK = 1024

            p = pyaudio.PyAudio()
            stream = p.open(
                format=pyaudio.paInt16,
                channels=1,
                rate=RATE,
                input=True,
                frames_per_buffer=CHUNK
            )

            audio_chunks = []

            while self.is_recording:
                try:
                    data = stream.read(CHUNK, exception_on_overflow=False)
                    audio_chunks.append(data)

                    # Calcular nivel de audio para o overlay
                    count = len(data) // 2
                    if count > 0:
                        shorts = struct.unpack(f"{count}h", data)
                        rms = math.sqrt(sum(s * s for s in shorts) / count)
                        level = min(1.0, rms / 5000)
                        set_audio_level(level)

                except Exception as e:
                    if not self.is_recording:
                        break

            stream.stop_stream()
            stream.close()
            p.terminate()

            set_audio_level(0.0)

            if audio_chunks:
                combined = b''.join(audio_chunks)
                self.audio_data = sr.AudioData(combined, RATE, 2)
                log.debug(f"Audio capturado: {len(combined)} bytes")
            else:
                log.warning("Nenhum audio capturado")
                self.audio_data = None

        except Exception as e:
            log.error(f"Erro na gravacao: {e}")
            set_audio_level(0.0)
            self.audio_data = None

    def transcribe(self):
        if not self.audio_data:
            log.warning("Sem audio para transcrever")
            play_sound('error')
            self.is_processing = False
            self.update_icon("ready")
            return

        log.info("Transcrevendo audio...")
        self.update_icon("processing")

        try:
            texto = self.recognizer.recognize_google(
                self.audio_data,
                language="pt-BR",
                show_all=True
            )

            if texto and 'alternative' in texto:
                texto = texto['alternative'][0]['transcript']
            elif isinstance(texto, str):
                pass
            else:
                texto = None

            if texto:
                log.info(f"Transcricao: '{texto}'")

                # Corrigir gagueira
                texto = corrigir_gagueira(texto)

                # Adicionar pontuacao
                texto = adicionar_pontuacao(texto)

                log.debug(f"Texto final: '{texto}'")

                # Adicionar ao historico
                self.historico.append(texto)
                if len(self.historico) > MAX_HISTORICO:
                    self.historico.pop(0)
                self.atualizar_menu()

                if self.paste_text(texto):
                    log.info("Texto colado com sucesso!")
                    play_sound('success')
            else:
                log.warning("Google nao retornou texto")
                play_sound('error')

        except sr.UnknownValueError:
            log.warning("Fala nao reconhecida")
            play_sound('error')
        except sr.RequestError as e:
            log.error(f"Erro de conexao com Google: {e}")
            play_sound('error')
        except Exception as e:
            log.error(f"Erro inesperado na transcricao: {e}")
            play_sound('error')
        finally:
            self.is_processing = False
            self.audio_data = None
            self.update_icon("ready")

    def on_press(self, event):
        if self.is_recording or self.is_processing:
            return

        self.target_window = self.get_active_window()
        self.is_recording = True
        self.record_start_time = time.time()
        play_sound('start')

        # Mostrar overlay
        try:
            get_overlay().mostrar()
        except:
            pass

        self.record_thread = threading.Thread(target=self.do_record, daemon=True)
        self.record_thread.start()

    def on_release(self, event):
        if not self.is_recording:
            return

        self.is_recording = False

        # Esconder overlay
        try:
            get_overlay().esconder()
        except:
            pass

        # Ignorar gravacoes muito curtas (< 0.3s) - evita acionamentos acidentais
        duracao = time.time() - self.record_start_time if self.record_start_time else 0
        if duracao < 0.3:
            self.update_icon("ready")
            return

        play_sound('stop')

        if hasattr(self, 'record_thread') and self.record_thread.is_alive():
            self.record_thread.join(timeout=2)

        self.is_processing = True
        threading.Thread(target=self.transcribe, daemon=True).start()

    def quit_app(self, icon=None, item=None):
        self.running = False
        keyboard.unhook_all()
        try:
            get_overlay().parar()
        except:
            pass
        if self.icon:
            self.icon.stop()

    def run(self):
        log.info("Registrando atalho de teclado...")

        # Limpar qualquer hook anterior
        keyboard.unhook_all()

        try:
            keyboard.on_press_key(TECLA, self.on_press, suppress=False)
            keyboard.on_release_key(TECLA, self.on_release, suppress=False)
            log.info(f"Atalho registrado: {TECLA_DISPLAY} ({TECLA})")
        except Exception as e:
            log.error(f"ERRO ao registrar atalho: {e}")
            log.error("Isso geralmente significa que precisa rodar como Administrador")
            _mostrar_erro_admin()
            raise

        menu = pystray.Menu(
            pystray.MenuItem(
                lambda item: f"✓ Colar automatico" if self.colar_automatico else "  Colar automatico",
                self.toggle_colar_automatico
            ),
            pystray.MenuItem("Configurar tecla", self.reconfigurar),
            pystray.MenuItem("─────────────", lambda: None, enabled=False),
            pystray.MenuItem("Como usar", self.mostrar_ajuda),
            pystray.MenuItem("Sair", self.quit_app)
        )

        self.icon = pystray.Icon(
            "VoiceToText",
            criar_icone("green"),
            "Segure CTRL DIREITO para gravar",
            menu
        )

        # Notificacao de boas-vindas
        def mostrar_notificacao():
            time.sleep(0.5)
            self.icon.notify("Segure CTRL DIREITO para gravar", "Voice to Text Ativo")
        threading.Thread(target=mostrar_notificacao, daemon=True).start()

        log.info("Voice to Text PRONTO!")
        log.info("-" * 50)
        self.icon.run()

    def abrir_log(self, icon=None, item=None):
        """Abre o arquivo de log"""
        try:
            os.startfile(LOG_FILE)
        except:
            pass

    def mostrar_ajuda(self, icon=None, item=None):
        """Mostra janela de ajuda"""
        import tkinter as tk

        root = tk.Tk()
        root.title("Voice to Text - Como Usar")
        root.geometry("380x300")
        root.resizable(False, False)
        root.attributes('-topmost', True)

        # Centralizar
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - 190
        y = (root.winfo_screenheight() // 2) - 150
        root.geometry(f"+{x}+{y}")

        texto = f"""
   COMO USAR

   1. Clique onde quer escrever

   2. SEGURE a tecla {TECLA_DISPLAY}

   3. Fale o que quer escrever

   4. SOLTE a tecla {TECLA_DISPLAY}

   5. Pronto!


   ICONES

   🟢 Verde = Pronto
   🔴 Vermelho = Gravando
   🟡 Amarelo = Processando
"""

        label = tk.Label(root, text=texto, font=("Consolas", 10),
                        justify="left", anchor="w")
        label.pack(padx=20, pady=10, fill="both", expand=True)

        tk.Button(root, text="OK", command=root.destroy,
                 font=("Segoe UI", 10), width=12).pack(pady=10)

        root.mainloop()

    def toggle_colar_automatico(self, icon=None, item=None):
        """Alterna entre colar automatico e apenas copiar"""
        self.colar_automatico = not self.colar_automatico

        # Salvar preferencia
        CONFIG["colar_automatico"] = self.colar_automatico
        salvar_config(CONFIG)

        # Notificar usuario
        if self.colar_automatico:
            self.icon.notify("Texto sera colado automaticamente", "Voice to Text")
        else:
            self.icon.notify("Texto sera apenas copiado", "Voice to Text")

        log.info(f"Colar automatico: {self.colar_automatico}")

    def reconfigurar(self, icon=None, item=None):
        """Abre janela para reconfigurar"""
        import tkinter as tk
        from tkinter import ttk, messagebox

        root = tk.Tk()
        root.title("Voice to Text - Configurar")
        root.geometry("380x200")
        root.resizable(False, False)
        root.attributes('-topmost', True)

        # Centralizar
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - 190
        y = (root.winfo_screenheight() // 2) - 100
        root.geometry(f"+{x}+{y}")

        tk.Label(root, text="Configurar Tecla",
                 font=("Segoe UI", 12, "bold")).pack(pady=15)

        frame = tk.Frame(root)
        frame.pack(pady=10, padx=20, fill="x")

        tk.Label(frame, text="Tecla para gravar:",
                 font=("Segoe UI", 10)).pack(anchor="w")

        # Encontrar indice atual
        atual = TECLA_DISPLAY
        tecla_var = tk.StringVar(value=atual)
        combo = ttk.Combobox(frame, textvariable=tecla_var,
                             values=[t[1] for t in TECLAS_DISPONIVEIS],
                             state="readonly", font=("Segoe UI", 10))
        combo.pack(fill="x", pady=5)

        def salvar():
            nome_exibicao = tecla_var.get()
            for interno, exibicao in TECLAS_DISPONIVEIS:
                if exibicao == nome_exibicao:
                    CONFIG["tecla"] = interno
                    CONFIG["tecla_display"] = exibicao
                    salvar_config(CONFIG)
                    break

            root.destroy()
            messagebox.showinfo(
                "Voice to Text",
                f"Tecla alterada para: {nome_exibicao}\n\n"
                "Reinicie o programa para aplicar."
            )

        tk.Button(root, text="Salvar", command=salvar,
                 font=("Segoe UI", 10), width=12).pack(pady=15)

        root.mainloop()


def main():
    log.info("Iniciando aplicacao...")

    # Verificar admin (aviso, nao obrigatorio)
    if not _verificar_admin():
        log.warning("NAO esta rodando como Administrador!")
        log.warning("Algumas funcoes podem nao funcionar corretamente")

    # Inicializar sons customizados
    log.debug("Inicializando sons...")
    inicializar_sons()

    try:
        app = VoiceToText()
        app.run()
    except KeyboardInterrupt:
        log.info("Encerrado pelo usuario")
    except Exception as e:
        log.error(f"ERRO FATAL: {e}")
        log.exception("Detalhes do erro:")
        play_sound('error')

        # Mostrar erro para o usuario
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Voice to Text - Erro",
            f"Ocorreu um erro:\n\n{str(e)}\n\n"
            f"Verifique o arquivo de log:\n{LOG_FILE}"
        )
        root.destroy()
        sys.exit(1)

    log.info("Aplicacao encerrada")


def verificar_instancia_unica():
    """Garante que só uma instancia do programa rode por vez"""
    import ctypes

    # Criar mutex para verificar instancia unica
    mutex_name = "VoiceToText_SingleInstance_Mutex"
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
    erro = ctypes.windll.kernel32.GetLastError()

    # ERROR_ALREADY_EXISTS = 183
    if erro == 183:
        log.warning("Outra instancia ja esta rodando!")

        # Fechar mutex
        ctypes.windll.kernel32.CloseHandle(mutex)

        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning(
            "Voice to Text",
            "O programa ja esta rodando!\n\n"
            "Verifique o icone na bandeja do sistema\n"
            "(perto do relogio)."
        )
        root.destroy()
        sys.exit(0)

    return mutex


if __name__ == "__main__":
    mutex = verificar_instancia_unica()
    main()

"""
Módulo de Criptografia para FAI-SQL
Usa AES-256 via Fernet com chave derivada de senha do usuário (PBKDF2)
"""

import os
import sys
import json
import base64
import hashlib
from cryptography.fernet import Fernet, InvalidToken
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC


def _get_base_dir():
    """Retorna o diretório base para os arquivos de dados.
    Se executando como .exe PyInstaller, retorna a pasta do .exe.
    Se executando como script Python, retorna a pasta do script."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


_BASE_DIR = _get_base_dir()

# Arquivos de dados criptografados (caminhos absolutos baseados na pasta do .exe)
ARQUIVO_CONFIG = os.path.join(_BASE_DIR, "config.dat")
ARQUIVO_COMANDOS = os.path.join(_BASE_DIR, "comandos.dat")
ARQUIVO_HISTORICO = os.path.join(_BASE_DIR, "historico.dat")
ARQUIVO_SALT = os.path.join(_BASE_DIR, ".salt")  # Arquivo que guarda o salt (indica que senha foi configurada)
ARQUIVO_VERIFICADOR = os.path.join(_BASE_DIR, ".verify")  # Para verificar se a senha está correta

# Variável global para armazenar a chave durante a sessão
_chave_sessao = None


def _gerar_salt():
    """Gera um salt aleatório de 16 bytes"""
    return os.urandom(16)


def _carregar_salt():
    """Carrega o salt salvo ou retorna None se não existir"""
    if os.path.exists(ARQUIVO_SALT):
        with open(ARQUIVO_SALT, "rb") as f:
            return f.read()
    return None


def _salvar_salt(salt):
    """Salva o salt no arquivo"""
    with open(ARQUIVO_SALT, "wb") as f:
        f.write(salt)


def _derivar_chave(senha, salt):
    """Deriva uma chave AES a partir da senha e salt usando PBKDF2"""
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=480000,  # Número alto para dificultar brute force
    )
    chave = base64.urlsafe_b64encode(kdf.derive(senha.encode()))
    return chave


def senha_configurada():
    """Verifica se já existe uma senha mestra configurada"""
    return os.path.exists(ARQUIVO_SALT) and os.path.exists(ARQUIVO_VERIFICADOR)


def configurar_senha(senha):
    """
    Configura a senha mestra pela primeira vez.
    Gera salt, deriva chave e salva verificador.
    Retorna True se sucesso.
    """
    global _chave_sessao
    
    if len(senha) < 4:
        return False, "Senha deve ter pelo menos 4 caracteres"
    
    # Gerar novo salt
    salt = _gerar_salt()
    _salvar_salt(salt)
    
    # Derivar chave
    chave = _derivar_chave(senha, salt)
    _chave_sessao = chave
    
    # Salvar verificador (texto conhecido criptografado para verificar senha)
    fernet = Fernet(chave)
    verificador = fernet.encrypt(b"FAISQL_VERIFICADOR_2024")
    with open(ARQUIVO_VERIFICADOR, "wb") as f:
        f.write(verificador)
    
    return True, "Senha configurada com sucesso"


def verificar_senha(senha):
    """
    Verifica se a senha informada está correta.
    Retorna True se correta, False caso contrário.
    """
    global _chave_sessao
    
    salt = _carregar_salt()
    if not salt:
        return False, "Senha não configurada"
    
    chave = _derivar_chave(senha, salt)
    
    try:
        with open(ARQUIVO_VERIFICADOR, "rb") as f:
            verificador_criptografado = f.read()
        
        fernet = Fernet(chave)
        texto = fernet.decrypt(verificador_criptografado)
        
        if texto == b"FAISQL_VERIFICADOR_2024":
            _chave_sessao = chave
            return True, "Senha correta"
        else:
            return False, "Senha incorreta"
    except InvalidToken:
        return False, "Senha incorreta"
    except Exception as e:
        return False, f"Erro: {str(e)}"


def resetar_dados():
    """
    Remove todos os dados criptografados e a senha.
    ATENÇÃO: Isso apaga tudo permanentemente!
    """
    global _chave_sessao
    _chave_sessao = None
    
    arquivos = [ARQUIVO_SALT, ARQUIVO_VERIFICADOR, 
                ARQUIVO_CONFIG, ARQUIVO_COMANDOS, ARQUIVO_HISTORICO]
    
    for arquivo in arquivos:
        if os.path.exists(arquivo):
            os.remove(arquivo)


def ler_arquivo_seguro(nome_arquivo):
    """
    Lê e descriptografa um arquivo de dados.
    Retorna o dicionário/lista ou dados padrão se arquivo não existir.
    """
    global _chave_sessao
    
    if _chave_sessao is None:
        raise Exception("Sessão não autenticada")
    
    if not os.path.exists(nome_arquivo):
        return None
    
    try:
        with open(nome_arquivo, "rb") as f:
            dados_criptografados = f.read()
        
        if not dados_criptografados:
            return None
        
        fernet = Fernet(_chave_sessao)
        dados_json = fernet.decrypt(dados_criptografados)
        return json.loads(dados_json.decode('utf-8'))
    except InvalidToken:
        raise Exception("Erro de descriptografia - dados corrompidos ou senha alterada")
    except json.JSONDecodeError:
        return None
    except Exception as e:
        raise Exception(f"Erro ao ler arquivo: {str(e)}")


def escrever_arquivo_seguro(nome_arquivo, dados):
    """
    Criptografa e salva dados em um arquivo.
    """
    global _chave_sessao
    
    if _chave_sessao is None:
        raise Exception("Sessão não autenticada")
    
    try:
        dados_json = json.dumps(dados, indent=2, ensure_ascii=False).encode('utf-8')
        fernet = Fernet(_chave_sessao)
        dados_criptografados = fernet.encrypt(dados_json)
        
        with open(nome_arquivo, "wb") as f:
            f.write(dados_criptografados)
    except Exception as e:
        raise Exception(f"Erro ao salvar arquivo: {str(e)}")


def migrar_dados_antigos():
    """
    Migra dados de arquivos JSON antigos (não criptografados) 
    para o formato criptografado.
    Retorna True se migrou algo.
    """
    migrou = False
    
    # Migrar config.json
    config_json = os.path.join(_BASE_DIR, "config.json")
    if os.path.exists(config_json) and not os.path.exists(ARQUIVO_CONFIG):
        try:
            with open(config_json, "r", encoding="utf-8") as f:
                dados = json.load(f)
            escrever_arquivo_seguro(ARQUIVO_CONFIG, dados)
            os.rename(config_json, config_json + ".backup")
            migrou = True
        except:
            pass
    
    # Migrar comandos.json
    comandos_json = os.path.join(_BASE_DIR, "comandos.json")
    if os.path.exists(comandos_json) and not os.path.exists(ARQUIVO_COMANDOS):
        try:
            with open(comandos_json, "r", encoding="utf-8") as f:
                dados = json.load(f)
            escrever_arquivo_seguro(ARQUIVO_COMANDOS, dados)
            os.rename(comandos_json, comandos_json + ".backup")
            migrou = True
        except:
            pass
    
    # Migrar historico.json
    historico_json = os.path.join(_BASE_DIR, "historico.json")
    if os.path.exists(historico_json) and not os.path.exists(ARQUIVO_HISTORICO):
        try:
            with open(historico_json, "r", encoding="utf-8") as f:
                dados = json.load(f)
            escrever_arquivo_seguro(ARQUIVO_HISTORICO, dados)
            os.rename(historico_json, historico_json + ".backup")
            migrou = True
        except:
            pass
    
    return migrou

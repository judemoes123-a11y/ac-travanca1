import os
import uuid
import math
import random
import sqlite3
from datetime import datetime
from flask import Flask, render_template, request, jsonify
from flask_socketio import SocketIO, emit, join_room
import pandas as pd

app = Flask(__name__, template_folder="templates" if os.path.exists("templates/index.html") else ".")
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'ac-travanca-secret-key-2026')

# Usamos async_mode='threading' com simple-websocket para máxima compatibilidade
# com qualquer versão do Python (3.10, 3.11, 3.12, 3.13, 3.14) sem erros do eventlet.
socketio = SocketIO(app, cors_allowed_origins="*", async_mode="threading")

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ranking_global.db")
EXCEL_JOGADORES = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Jogadores e Estatísticas.xlsx")
EXCEL_EQUIPAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Lista de equipas e Força.xlsx")


# ==============================================================================
# 1. BASE DE DADOS DO TOP 10 GLOBAL DE TODOS OS UTILIZADORES (MODO OFFLINE)
# ==============================================================================
def init_db():
    """Inicializa a base de dados SQLite que guarda o histórico de TODOS os utilizadores
    que já jogaram o modo offline para construir o Top 10 Global."""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS historico_utilizadores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_utilizador TEXT NOT NULL,
            pontos INTEGER NOT NULL,
            saldo_golos INTEGER NOT NULL,
            golos_marcados INTEGER NOT NULL,
            posicao_liga INTEGER NOT NULL,
            tatica INTEGER NOT NULL,
            data_registo TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()


def registar_pontuacao_global(nome_utilizador, pontos, saldo_golos, golos_marcados, posicao_liga, tatica):
    """Guarda o resultado do utilizador no histórico global e devolve:
    - O Top 10 de todos os utilizadores que já jogaram o jogo
    - A posição exata (ex: 16º) em que esta participação do utilizador ficou no ranking global
    - O total de participações registadas no jogo
    """
    nome_limpo = (nome_utilizador or "Treinador Anónimo").strip()[:40]
    agora = datetime.now().strftime("%Y-%m-%d %H:%M")

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        INSERT INTO historico_utilizadores
        (nome_utilizador, pontos, saldo_golos, golos_marcados, posicao_liga, tatica, data_registo)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (nome_limpo, int(pontos), int(saldo_golos), int(golos_marcados), int(posicao_liga), int(tatica), agora))
    run_id = c.lastrowid
    conn.commit()

    # Ordena TODOS os utilizadores que já jogaram por Pontos DESC, Saldo de Golos DESC, Golos Marcados DESC
    c.execute("""
        SELECT id, nome_utilizador, pontos, saldo_golos, golos_marcados, posicao_liga, tatica, data_registo
        FROM historico_utilizadores
        ORDER BY pontos DESC, saldo_golos DESC, golos_marcados DESC, id ASC
    """)
    todas_linhas = c.fetchall()
    conn.close()

    top_10 = []
    minha_posicao_global = None
    meu_registo = None

    for idx, row in enumerate(todas_linhas, start=1):
        item = {
            "lugar_global": idx,
            "id": row[0],
            "nome_utilizador": row[1],
            "pontos": row[2],
            "saldo_golos": row[3],
            "golos_marcados": row[4],
            "posicao_liga": row[5],
            "tatica": row[6],
            "data_registo": row[7],
            "atual": (row[0] == run_id)
        }
        if idx <= 10:
            top_10.append(item)
        if row[0] == run_id:
            minha_posicao_global = idx
            meu_registo = item

    return {
        "top_10": top_10,
        "minha_posicao_global": minha_posicao_global,
        "meu_registo": meu_registo,
        "total_jogadores_historico": len(todas_linhas)
    }


def obter_top_10_global():
    """Retorna o Top 10 Global atual de todos os utilizadores que já jogaram."""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        SELECT id, nome_utilizador, pontos, saldo_golos, golos_marcados, posicao_liga, tatica, data_registo
        FROM historico_utilizadores
        ORDER BY pontos DESC, saldo_golos DESC, golos_marcados DESC, id ASC
        LIMIT 10
    """)
    rows = c.fetchall()
    c.execute("SELECT COUNT(*) FROM historico_utilizadores")
    total = c.fetchone()[0]
    conn.close()

    top_10 = []
    for idx, row in enumerate(rows, start=1):
        top_10.append({
            "lugar_global": idx,
            "nome_utilizador": row[1],
            "pontos": row[2],
            "saldo_golos": row[3],
            "golos_marcados": row[4],
            "posicao_liga": row[5],
            "tatica": row[6],
            "data_registo": row[7]
        })
    return {"top_10": top_10, "total_jogadores_historico": total}


# ==============================================================================
# 2. LEITURA DINÂMICA DOS FICHEIROS EXCEL COM PANDAS
# ==============================================================================
def carregar_jogadores():
    """Lê dinamicamente o ficheiro Excel 'Jogadores e Estatísticas.xlsx' com pandas.
    Junta guarda-redes e jogadores de campo numa lista única para o sorteio."""
    df = pd.read_excel(EXCEL_JOGADORES)
    jogadores = []
    idx_id = 1

    # Tabela da esquerda: Guarda-Redes
    if 'Nome' in df.columns and 'Habilidade como Guarda-Redes' in df.columns:
        for _, row in df.iterrows():
            if pd.notna(row.get('Nome')) and pd.notna(row.get('Habilidade como Guarda-Redes')):
                jogadores.append({
                    'id': idx_id,
                    'nome': str(row['Nome']).strip(),
                    'posicao': str(row['Posição']).strip(),
                    'eh_gr_natural': True,
                    'gr': float(row['Habilidade como Guarda-Redes']),
                    'defesa': 40.0,
                    'drible': 40.0,
                    'remate': 40.0,
                    'fisico': 40.0,
                    'passe': 40.0,
                    'velocidade': 40.0
                })
                idx_id += 1

    # Tabela da direita: Jogadores de Linha
    if 'Nome.1' in df.columns:
        for _, row in df.iterrows():
            if pd.notna(row.get('Nome.1')):
                pos_str = str(row['Posição.1']).strip()
                jogadores.append({
                    'id': idx_id,
                    'nome': str(row['Nome.1']).strip(),
                    'posicao': pos_str,
                    'eh_gr_natural': ('guarda-redes' in pos_str.lower()),
                    'gr': 40.0,
                    'defesa': float(row['Habilidade Defesa']),
                    'drible': float(row['Habilidade Drible']),
                    'remate': float(row['Habilidade Remate']),
                    'fisico': float(row['Habilidade Físico']),
                    'passe': float(row['Habilidade Passe']),
                    'velocidade': float(row['Habilidade Velocidade'])
                })
                idx_id += 1

    return jogadores


def carregar_equipas():
    """Lê dinamicamente o ficheiro Excel 'Lista de equipas e Força.xlsx' com pandas."""
    df = pd.read_excel(EXCEL_EQUIPAS)
    equipas = []
    for _, row in df.iterrows():
        if pd.notna(row.get('Nome da Equipa')) and pd.notna(row.get('Força da Equipa')):
            equipas.append({
                'nome': str(row['Nome da Equipa']).strip(),
                'forca': float(row['Força da Equipa'])
            })
    return equipas


def sanitizar_jogador_publico(j):
    """REGRA CRÍTICA: Esconde todas as estatísticas dos jogadores.
    O utilizador em momento algum sabe os atributos numéricos."""
    return {
        "id": j["id"],
        "nome": j["nome"],
        "posicao": j["posicao"]
    }


# ==============================================================================
# 3. ESTRUTURA DE CLASSES E REGRAS DE POSICIONAMENTO / TÁTICAS
# ==============================================================================
# Mapeamento de cada slot (1 a 7) em cada tática (1, 2, 3) para o setor correspondente
MAPA_SLOTS_TATICA = {
    1: {  # Tática 1: 1 GR, 2 Defesas (2,3), 3 Médios (4 central; 5,6 laterais), 1 Avançado (7)
        1: ("Guarda-redes", "GR"),
        2: ("Defesa", "DEF_T1"),
        3: ("Defesa", "DEF_T1"),
        4: ("Médio", "MED_CENTRAL_T1"),
        5: ("Médio", "MED_LATERAL_T1"),
        6: ("Médio", "MED_LATERAL_T1"),
        7: ("Avançado", "AVA_T1")
    },
    2: {  # Tática 2: 1 GR, 3 Defesas (4 central; 2,3 laterais), 1 Médio (5), 2 Avançados (6,7)
        1: ("Guarda-redes", "GR"),
        2: ("Defesa", "DEF_LATERAL_T2"),
        3: ("Defesa", "DEF_LATERAL_T2"),
        4: ("Defesa", "DEF_CENTRAL_T2"),
        5: ("Médio", "MED_T2"),
        6: ("Avançado", "AVA_T2"),
        7: ("Avançado", "AVA_T2")
    },
    3: {  # Tática 3: 1 GR, 2 Defesas (2,4), 2 Médios (3,5), 2 Avançados (6,7)
        1: ("Guarda-redes", "GR"),
        2: ("Defesa", "DEF_T3"),
        3: ("Médio", "MED_T3"),
        4: ("Defesa", "DEF_T3"),
        5: ("Médio", "MED_T3"),
        6: ("Avançado", "AVA_T3"),
        7: ("Avançado", "AVA_T3")
    }
}


def jogador_esta_na_posicao_natural(pos_natural_str, setor_slot):
    """Verifica se o setor atribuído (Guarda-redes, Defesa, Médio, Avançado)
    faz parte das posições naturais do jogador (ex: 'Defesa/Médio')."""
    partes = [p.strip().lower() for p in pos_natural_str.split('/')]
    return setor_slot.strip().lower() in partes


class GuardaRedes:
    """Classe que representa o Guarda-Redes (Atributo Único: Habilidade como GR)."""
    def __init__(self, dados_jogador, slot_num=1):
        self.id = dados_jogador['id']
        self.nome = dados_jogador['nome']
        self.posicao_natural = dados_jogador['posicao']
        self.slot_num = slot_num
        self.eh_gr_natural = dados_jogador['eh_gr_natural']

        # Regra: Se o jogador selecionado a guarda-redes não for dessa posição, fica com 40
        if self.eh_gr_natural:
            self.habilidade_gr = float(dados_jogador['gr'])
        else:
            self.habilidade_gr = 40.0

        self.fadiga = 0.0  # GR sofre desgaste mínimo


class JogadorDeLinha:
    """Classe que representa um Jogador de Linha (6 atributos + Fadiga Dinâmica)."""
    def __init__(self, dados_jogador, setor_atribuido, codigo_calculo, slot_num):
        self.id = dados_jogador['id']
        self.nome = dados_jogador['nome']
        self.posicao_natural = dados_jogador['posicao']
        self.setor_atribuido = setor_atribuido  # 'Defesa', 'Médio' ou 'Avançado'
        self.codigo_calculo = codigo_calculo
        self.slot_num = slot_num
        self.eh_gr_natural = dados_jogador['eh_gr_natural']

        # Regra 1: Se um Guarda-Redes estiver a jogar noutra posição, fica com 40 a tudo
        if self.eh_gr_natural:
            self.defesa_base = 40.0
            self.drible_base = 40.0
            self.remate_base = 40.0
            self.fisico_base = 40.0
            self.passe_base = 40.0
            self.velocidade_base = 40.0
            self.fora_de_posicao = True
        else:
            # Regra 2: Se estiver numa posição que não é a dele, todas as estatísticas baixam -5
            na_pos = jogador_esta_na_posicao_natural(self.posicao_natural, self.setor_atribuido)
            penalizacao = 0.0 if na_pos else 5.0
            self.fora_de_posicao = not na_pos

            self.defesa_base = max(1.0, float(dados_jogador['defesa']) - penalizacao)
            self.drible_base = max(1.0, float(dados_jogador['drible']) - penalizacao)
            self.remate_base = max(1.0, float(dados_jogador['remate']) - penalizacao)
            self.fisico_base = max(1.0, float(dados_jogador['fisico']) - penalizacao)
            self.passe_base = max(1.0, float(dados_jogador['passe']) - penalizacao)
            self.velocidade_base = max(1.0, float(dados_jogador['velocidade']) - penalizacao)

        # Estado dinâmico para o motor de jogo online (Stamina / Fadiga)
        self.stamina = 100.0  # Começa a 100%

    def fator_fadiga(self):
        """Calcula o multiplicador de performance atual com base na stamina restante.
        Abaixo de 85% de stamina, os atributos técnicos e físicos sofrem penalização progressiva."""
        return max(0.55, 0.55 + 0.45 * (self.stamina / 100.0))

    @property
    def defesa(self):
        return self.defesa_base * self.fator_fadiga()

    @property
    def drible(self):
        return self.drible_base * self.fator_fadiga()

    @property
    def remate(self):
        return self.remate_base * self.fator_fadiga()

    @property
    def fisico(self):
        return self.fisico_base * self.fator_fadiga()

    @property
    def passe(self):
        return self.passe_base * self.fator_fadiga()

    @property
    def velocidade(self):
        return self.velocidade_base * self.fator_fadiga()

    def aplicar_desgaste(self, intensidade=1.0):
        """Fadiga Dinâmica (Stamina Decay):
        O atributo Físico base dita a resistência. Jogadores com Físico baixo perdem
        stamina muito mais depressa a cada minuto/ação."""
        resistencia = self.fisico_base / 100.0  # Ex: 90 Físico -> 0.90; 60 Físico -> 0.60
        perda = intensidade * (1.45 - 0.85 * resistencia)
        self.stamina = max(15.0, self.stamina - perda)


# ==============================================================================
# 4. CÁLCULO DE FORÇA E SIMULAÇÃO DA LIGA OFFLINE (simular_liga_offline)
# ==============================================================================
def calcular_forca_jogador_offline(dados_jogador, tatica, slot_num):
    """Calcula a força individual de um jogador no Modo Offline segundo a tática (1, 2 ou 3)
    e a posição (slot 1 a 7) em que foi colocado."""
    setor, codigo = MAPA_SLOTS_TATICA[int(tatica)][int(slot_num)]

    if setor == "Guarda-redes":
        gr_obj = GuardaRedes(dados_jogador, slot_num)
        return gr_obj.habilidade_gr

    jog = JogadorDeLinha(dados_jogador, setor, codigo, slot_num)
    D = jog.defesa_base
    Dr = jog.drible_base
    R = jog.remate_base
    F = jog.fisico_base
    P = jog.passe_base
    V = jog.velocidade_base

    # Média dos 3 atributos agrupados quando a fórmula indica Drible+Remate+Velocidade(X%)
    media_drv = (Dr + R + V) / 3.0

    if codigo == "DEF_T1":
        # Defesa (42%), Físico (42%), Passe (11%), Drible+Remate+Velocidade (5%)
        return D * 0.42 + F * 0.42 + P * 0.11 + media_drv * 0.05
    elif codigo == "MED_CENTRAL_T1":
        # Defesa (10%), Drible (30%), Físico (10%), Passe (35%), Velocidade (5%), Remate (10%)
        return D * 0.10 + Dr * 0.30 + F * 0.10 + P * 0.35 + V * 0.05 + R * 0.10
    elif codigo == "MED_LATERAL_T1":
        # Drible (30%), Remate (15%), Passe (25%), Velocidade (20%), Defesa (5%), Físico (5%)
        return Dr * 0.30 + R * 0.15 + P * 0.25 + V * 0.20 + D * 0.05 + F * 0.05
    elif codigo in ("AVA_T1", "AVA_T2", "AVA_T3"):
        # Drible (23%), Remate (28%), Físico (23%), Passe (10%), Velocidade (15%), Defesa (1%)
        return Dr * 0.23 + R * 0.28 + F * 0.23 + P * 0.10 + V * 0.15 + D * 0.01
    elif codigo == "DEF_CENTRAL_T2":
        # Defesa (43%), Físico (43%), Passe (10%), Drible+Remate+Velocidade (4%)
        return D * 0.43 + F * 0.43 + P * 0.10 + media_drv * 0.04
    elif codigo == "DEF_LATERAL_T2":
        # Defesa (31%), Físico (31%), Passe (15%), Velocidade (10%), Drible (10%), Remate (3%)
        return D * 0.31 + F * 0.31 + P * 0.15 + V * 0.10 + Dr * 0.10 + R * 0.03
    elif codigo == "MED_T2":
        # Defesa (10%), Drible (30%), Físico (10%), Passe (35%), Velocidade (5%), Remate (10%)
        return D * 0.10 + Dr * 0.30 + F * 0.10 + P * 0.35 + V * 0.05 + R * 0.10
    elif codigo == "DEF_T3":
        # Defesa (42%), Físico (42%), Passe (11%), Drible+Remate+Velocidade (5%)
        return D * 0.42 + F * 0.42 + P * 0.11 + media_drv * 0.05
    elif codigo == "MED_T3":
        # Defesa (10%), Drible (30%), Físico (10%), Passe (35%), Velocidade (10%), Remate (5%)
        return D * 0.10 + Dr * 0.30 + F * 0.10 + P * 0.35 + V * 0.10 + R * 0.05

    return (D + Dr + R + F + P + V) / 6.0


def simular_jogo_rapido(forca_casa, forca_fora):
    """Simula o resultado de um jogo no Modo Offline tendo em conta principalmente
    a diferença de força entre as equipas e uma aleatoriedade baixa."""
    diff = forca_casa - forca_fora
    # Ruído gaussiano baixo para permitir pequenos imprevistos sem distorcer a força real
    desempenho_casa = forca_casa + random.gauss(0, 1.8)
    desempenho_fora = forca_fora + random.gauss(0, 1.8)
    vantagem = desempenho_casa - desempenho_fora

    # Expectativa de golos (xG) ajustada pela diferença de força
    xg_casa = max(0.3, 1.65 + vantagem * 0.22)
    xg_fora = max(0.3, 1.50 - vantagem * 0.22)

    golos_casa = max(0, int(round(random.gauss(xg_casa, 0.75))))
    golos_fora = max(0, int(round(random.gauss(xg_fora, 0.75))))
    return golos_casa, golos_fora


def simular_liga_offline(alinhamento_slots, tatica, nome_utilizador="Treinador"):
    """Executa a simulação completa da Liga Offline (2 voltas = 2 jogos contra cada equipa),
    regista os marcadores dos golos da equipa do utilizador e guarda o resultado no
    Ranking Global (Top 10 de todos os utilizadores que já jogaram)."""
    todos_jogadores = {j['id']: j for j in carregar_jogadores()}
    equipas_ia = carregar_equipas()

    # 1. Calcula a força média dos 7 jogadores do utilizador
    forcas_individuais = []
    pesos_golo_jogadores = []

    for slot_str, jog_id in alinhamento_slots.items():
        slot_num = int(slot_str)
        j_dados = todos_jogadores[int(jog_id)]
        f_jog = calcular_forca_jogador_offline(j_dados, tatica, slot_num)
        forcas_individuais.append(f_jog)

        setor, codigo = MAPA_SLOTS_TATICA[int(tatica)][slot_num]
        if setor == "Guarda-redes":
            peso_golo = 0.01
        else:
            jog_obj = JogadorDeLinha(j_dados, setor, codigo, slot_num)
            mult_setor = 3.2 if setor == "Avançado" else (1.7 if setor == "Médio" else 0.6)
            peso_golo = ((jog_obj.remate_base ** 2) * 0.65 + (jog_obj.drible_base ** 2) * 0.35) * mult_setor
        pesos_golo_jogadores.append((j_dados['nome'], j_dados['posicao'], setor, peso_golo))

    forca_utilizador = sum(forcas_individuais) / 7.0

    # 2. Prepara lista de todas as equipas da Liga (AC Travanca + 11 adversários)
    nome_equipa_user = "AC Travanca (Tu)"
    participantes = [{"nome": nome_equipa_user, "forca": forca_utilizador, "eh_user": True}]
    for eq in equipas_ia:
        participantes.append({"nome": eq["nome"], "forca": eq["forca"], "eh_user": False})

    tabela = {
        p["nome"]: {
            "equipa": p["nome"],
            "eh_user": p["eh_user"],
            "jogos": 0,
            "vitorias": 0,
            "empates": 0,
            "derrotas": 0,
            "golos_marcados": 0,
            "golos_sofridos": 0,
            "saldo_golos": 0,
            "pontos": 0
        } for p in participantes
    }

    marcadores_user = {item[0]: {"nome": item[0], "posicao": item[2], "golos": 0} for item in pesos_golo_jogadores}
    nomes_sorteio_golo = [item[0] for item in pesos_golo_jogadores]
    pesos_sorteio_golo = [item[3] for item in pesos_golo_jogadores]

    historico_jogos_user = []

    # 3. Simulação das 2 Voltas (Casa e Fora entre todos os pares de equipas)
    for volta in (1, 2):
        for i in range(len(participantes)):
            for j in range(i + 1, len(participantes)):
                eq_casa = participantes[i] if volta == 1 else participantes[j]
                eq_fora = participantes[j] if volta == 1 else participantes[i]

                gc, gf = simular_jogo_rapido(eq_casa["forca"], eq_fora["forca"])

                # Atualiza estatísticas da equipa da casa
                tc = tabela[eq_casa["nome"]]
                tf = tabela[eq_fora["nome"]]
                tc["jogos"] += 1
                tf["jogos"] += 1
                tc["golos_marcados"] += gc
                tc["golos_sofridos"] += gf
                tc["saldo_golos"] = tc["golos_marcados"] - tc["golos_sofridos"]

                tf["golos_marcados"] += gf
                tf["golos_sofridos"] += gc
                tf["saldo_golos"] = tf["golos_marcados"] - tf["golos_sofridos"]

                if gc > gf:
                    tc["vitorias"] += 1
                    tc["pontos"] += 3
                    tf["derrotas"] += 1
                elif gc < gf:
                    tf["vitorias"] += 1
                    tf["pontos"] += 3
                    tc["derrotas"] += 1
                else:
                    tc["empates"] += 1
                    tf["empates"] += 1
                    tc["pontos"] += 1
                    tf["pontos"] += 1

                # Se o utilizador participou neste jogo, atribui os marcadores dos seus golos
                if eq_casa["eh_user"] or eq_fora["eh_user"]:
                    golos_user = gc if eq_casa["eh_user"] else gf
                    autores = []
                    for _ in range(golos_user):
                        autor = random.choices(nomes_sorteio_golo, weights=pesos_sorteio_golo, k=1)[0]
                        marcadores_user[autor]["golos"] += 1
                        autores.append(autor)

                    historico_jogos_user.append({
                        "volta": volta,
                        "casa": eq_casa["nome"],
                        "fora": eq_fora["nome"],
                        "golos_casa": gc,
                        "golos_fora": gf,
                        "marcadores_user": autores
                    })

    # 4. Ordena Tabela Classificativa Final por Pontos DESC, Saldo de Golos DESC, Golos Marcados DESC
    classificacao = sorted(
        tabela.values(),
        key=lambda x: (x["pontos"], x["saldo_golos"], x["golos_marcados"]),
        reverse=True
    )
    posicao_final_liga = 1
    for idx, linha in enumerate(classificacao, start=1):
        linha["lugar"] = idx
        if linha["eh_user"]:
            posicao_final_liga = idx

    lista_marcadores = sorted(
        [m for m in marcadores_user.values() if m["golos"] > 0],
        key=lambda x: x["golos"],
        reverse=True
    )

    dados_user_liga = tabela[nome_equipa_user]

    # 5. Regista a pontuação no Ranking Global Histórico de Todos os Utilizadores
    ranking_global = registar_pontuacao_global(
        nome_utilizador=nome_utilizador,
        pontos=dados_user_liga["pontos"],
        saldo_golos=dados_user_liga["saldo_golos"],
        golos_marcados=dados_user_liga["golos_marcados"],
        posicao_liga=posicao_final_liga,
        tatica=int(tatica)
    )

    return {
        "classificacao_liga": classificacao,
        "melhores_marcadores": lista_marcadores,
        "jogos_utilizador": historico_jogos_user,
        "resumo_utilizador": {
            "nome_utilizador": nome_utilizador,
            "lugar_liga": posicao_final_liga,
            "pontos": dados_user_liga["pontos"],
            "saldo_golos": dados_user_liga["saldo_golos"],
            "golos_marcados": dados_user_liga["golos_marcados"],
            "golos_sofridos": dados_user_liga["golos_sofridos"]
        },
        "ranking_global_utilizadores": ranking_global
    }


# ==============================================================================
# 5. MOTOR DE SIMULAÇÃO ONLINE DE ALTA COMPLEXIDADE (match_engine_online)
# ==============================================================================
class MatchEngine:
    """Motor de Jogo de Futebol de 7 ao estilo Football Manager.
    Baseado estritamente em confrontos matemáticos de atributos, zonas de campo,
    árvore de decisão probabilística, fadiga dinâmica e distribuições Gaussianas (curva de sino)."""

    def __init__(self, nome_eq_a, alinhamento_a, tatica_a, nome_eq_b, alinhamento_b, tatica_b):
        todos_jogadores = {j['id']: j for j in carregar_jogadores()}
        self.nome_a = nome_eq_a
        self.nome_b = nome_eq_b
        self.gr_a, self.linha_a = self._construir_equipa(alinhamento_a, tatica_a, todos_jogadores)
        self.gr_b, self.linha_b = self._construir_equipa(alinhamento_b, tatica_b, todos_jogadores)

        self.narracao = []
        self.stats = {
            "A": {
                "nome": nome_eq_a, "golos": 0, "ticks_posse": 0,
                "remates_total": 0, "remates_baliza": 0, "defesas_gr": 0,
                "passes_tentados": 0, "passes_completos": 0,
                "desarmes": 0, "intercecoes": 0
            },
            "B": {
                "nome": nome_eq_b, "golos": 0, "ticks_posse": 0,
                "remates_total": 0, "remates_baliza": 0, "defesas_gr": 0,
                "passes_tentados": 0, "passes_completos": 0,
                "desarmes": 0, "intercecoes": 0
            }
        }

    def _construir_equipa(self, alinhamento_slots, tatica, todos_jogadores):
        gr = None
        linha = []
        for slot_str, jog_id in alinhamento_slots.items():
            slot_num = int(slot_str)
            dados = todos_jogadores[int(jog_id)]
            setor, codigo = MAPA_SLOTS_TATICA[int(tatica)][slot_num]
            if setor == "Guarda-redes":
                gr = GuardaRedes(dados, slot_num)
            else:
                linha.append(JogadorDeLinha(dados, setor, codigo, slot_num))
        return gr, linha

    @staticmethod
    def _gauss_roll(media, desvio=9.5):
        """RNG Avançado (Curva de Sino / Gaussiana):
        Usa random.gauss(mu, sigma) onde o valor composto dos atributos é a média (mu).
        Garante consistência aos melhores jogadores mantendo probabilidade realista de surpresa."""
        return random.gauss(media, desvio)

    def _escolher_portador_inicial(self, equipa_linha, zona):
        """Seleciona um jogador da equipa para conduzir a bola consoante a zona atual."""
        pesos = []
        for j in equipa_linha:
            if zona == "Defesa":
                p = 3.0 if j.setor_atribuido == "Defesa" else (1.8 if j.setor_atribuido == "Médio" else 0.5)
            elif zona == "Meio-Campo":
                p = 3.2 if j.setor_atribuido == "Médio" else 1.5
            else:  # Ataque
                p = 3.5 if j.setor_atribuido == "Avançado" else (2.0 if j.setor_atribuido == "Médio" else 0.6)
            pesos.append(p)
        return random.choices(equipa_linha, weights=pesos, k=1)[0]

    def _escolher_defensor(self, equipa_defensora, zona_ataque):
        """Escolhe o opositor direto dependendo da zona onde decorre a jogada."""
        pesos = []
        for j in equipa_defensora:
            if zona_ataque == "Ataque":  # Zona defensiva do adversário
                p = 3.5 if j.setor_atribuido == "Defesa" else (1.6 if j.setor_atribuido == "Médio" else 0.4)
            elif zona_ataque == "Meio-Campo":
                p = 3.2 if j.setor_atribuido == "Médio" else (1.8 if j.setor_atribuido == "Defesa" else 1.0)
            else:  # Pressão alta
                p = 3.0 if j.setor_atribuido == "Avançado" else (2.0 if j.setor_atribuido == "Médio" else 0.7)
            pesos.append(p)
        return random.choices(equipa_defensora, weights=pesos, k=1)[0]

    def _decidir_acao_ia(self, portador, zona):
        """Árvore de Decisão Baseada em Atributos (IA do Jogador):
        Calcula a probabilidade de Passar, Driblar ou Rematar segundo os atributos
        do próprio jogador e a zona do campo em que se encontra."""
        if zona == "Defesa":
            # Na defesa privilegia a saída de bola segura (Passe) ou progressão (Drible)
            peso_passe = (portador.passe ** 1.35) * 1.4
            peso_drible = (portador.drible ** 1.25) * 0.45
            peso_remate = 0.0  # Não remata desde a própria defesa
        elif zona == "Meio-Campo":
            # No meio-campo pode construir (Passe), romper linhas (Drible) ou tentar meia-distância (Remate)
            peso_passe = (portador.passe ** 1.30) * 1.15
            peso_drible = (portador.drible ** 1.30) * 0.85
            peso_remate = (portador.remate ** 1.35) * 0.28
        else:  # Zona de Ataque
            peso_passe = (portador.passe ** 1.20) * 0.65
            peso_drible = (portador.drible ** 1.25) * 0.80
            peso_remate = (portador.remate ** 1.40) * 1.45

        acoes = ["PASSAR", "DRIBLAR", "REMATAR"]
        return random.choices(acoes, weights=[peso_passe, peso_drible, peso_remate], k=1)[0]

    def simular_partida(self):
        """Simula 60 minutos de Futebol de 7 com eventos dinâmicos minuto a minuto."""
        posse = "A" if random.random() < 0.5 else "B"
        zona = "Meio-Campo"

        self.narracao.append({
            "minuto": 0,
            "tipo": "INFO",
            "texto": f"Apito inicial! {self.nome_a} e {self.nome_b} já disputam a partida de Futebol de 7!"
        })

        for minuto in range(1, 61):
            if minuto == 31:
                self.narracao.append({
                    "minuto": 30,
                    "tipo": "INFO",
                    "texto": f"Intervalo! Resultado parcial: {self.nome_a} {self.stats['A']['golos']} - {self.stats['B']['golos']} {self.nome_b}."
                })

            eq_atk = self.linha_a if posse == "A" else self.linha_b
            eq_def = self.linha_b if posse == "A" else self.linha_a
            gr_def = self.gr_b if posse == "A" else self.gr_a
            st_atk = self.stats[posse]
            chave_def = "B" if posse == "A" else "A"
            st_def = self.stats[chave_def]
            nome_atk = self.nome_a if posse == "A" else self.nome_b
            nome_def = self.nome_b if posse == "A" else self.nome_a

            st_atk["ticks_posse"] += 1

            portador = self._escolher_portador_inicial(eq_atk, zona)
            marcador = self._escolher_defensor(eq_def, zona)
            acao = self._decidir_acao_ia(portador, zona)

            # Desgaste físico base do minuto para todos os jogadores em campo
            for j in self.linha_a + self.linha_b:
                j.aplicar_desgaste(intensidade=0.55)

            if acao == "PASSAR":
                st_atk["passes_tentados"] += 1
                portador.aplicar_desgaste(0.35)
                marcador.aplicar_desgaste(0.35)

                # Pressão defensiva aumenta consoante a zona (Ataque tem mais pressão que Defesa)
                pressao_zona = 1.08 if zona == "Ataque" else (1.00 if zona == "Meio-Campo" else 0.88)

                # Fórmula Ponderada de Passe vs Interceção:
                # Passe: (Passe * 0.75 + Drible * 0.25) vs Interceção: (Defesa * 0.70 + Velocidade * 0.30) * pressao_zona
                mu_passe = portador.passe * 0.75 + portador.drible * 0.25
                mu_intercecao = (marcador.defesa * 0.70 + marcador.velocidade * 0.30) * pressao_zona

                val_passe = self._gauss_roll(mu_passe, 8.5)
                val_inter = self._gauss_roll(mu_intercecao, 8.5)

                if val_passe >= val_inter - 4.0:
                    st_atk["passes_completos"] += 1
                    colegas = [c for c in eq_atk if c.id != portador.id]
                    recetor = random.choice(colegas) if colegas else portador
                    if zona == "Defesa":
                        zona = "Meio-Campo"
                        self.narracao.append({
                            "minuto": minuto, "tipo": "JOGADA",
                            "texto": f"[{minuto}'] {portador.nome} ({nome_atk}) sai a jogar com critério desde a defesa e entrega em {recetor.nome} no meio-campo."
                        })
                    elif zona == "Meio-Campo":
                        zona = "Ataque"
                        self.narracao.append({
                            "minuto": minuto, "tipo": "JOGADA",
                            "texto": f"[{minuto}'] Grande visão de jogo de {portador.nome} ({nome_atk})! Passe rasgado a encontrar {recetor.nome} na zona de ataque."
                        })
                    else:
                        # Passe de rutura na zona de ataque -> cria oportunidade imediata de finalização!
                        self._resolver_remate(minuto, recetor, marcador, gr_def, "Ataque", st_atk, st_def, nome_atk, nome_def, assistente=portador)
                        posse = chave_def
                        zona = "Defesa"
                else:
                    st_def["intercecoes"] += 1
                    posse = chave_def
                    zona = "Meio-Campo"
                    self.narracao.append({
                        "minuto": minuto, "tipo": "DEFESA",
                        "texto": f"[{minuto}'] Interceção imperial de {marcador.nome} ({nome_def}), que leu o passe de {portador.nome} e recupera a posse!"
                    })

            elif acao == "DRIBLAR":
                portador.aplicar_desgaste(0.65)
                marcador.aplicar_desgaste(0.60)

                # Fórmula Ponderada de Duelo 1v1:
                # Atacante: (Drible * 0.70 + Velocidade * 0.30) vs Defesa: (Defesa * 0.60 + Físico * 0.40)
                mu_atk = portador.drible * 0.70 + portador.velocidade * 0.30
                mu_def = marcador.defesa * 0.60 + marcador.fisico * 0.40

                val_atk = self._gauss_roll(mu_atk, 9.0)
                val_def = self._gauss_roll(mu_def, 9.0)

                if val_atk > val_def:
                    if zona == "Defesa":
                        zona = "Meio-Campo"
                        self.narracao.append({
                            "minuto": minuto, "tipo": "JOGADA",
                            "texto": f"[{minuto}'] {portador.nome} ({nome_atk}) arranca em velocidade, supera {marcador.nome} e galga metros até ao meio-campo!"
                        })
                    elif zona == "Meio-Campo":
                        zona = "Ataque"
                        self.narracao.append({
                            "minuto": minuto, "tipo": "JOGADA",
                            "texto": f"[{minuto}'] Que pormenor técnico! {portador.nome} ({nome_atk}) tira {marcador.nome} do caminho no 1v1 e entra na área adversária!"
                        })
                    else:
                        # Drible ganho no ataque deixa o jogador na cara do GR!
                        self.narracao.append({
                            "minuto": minuto, "tipo": "PERIGO",
                            "texto": f"[{minuto}'] {portador.nome} ({nome_atk}) finta {marcador.nome} na área e fica isolado perante {gr_def.nome}!"
                        })
                        self._resolver_remate(minuto, portador, marcador, gr_def, "Cara-a-Cara", st_atk, st_def, nome_atk, nome_def)
                        posse = chave_def
                        zona = "Defesa"
                else:
                    st_def["desarmes"] += 1
                    posse = chave_def
                    zona = "Meio-Campo"
                    self.narracao.append({
                        "minuto": minuto, "tipo": "DEFESA",
                        "texto": f"[{minuto}'] Desarme limpo e vigoroso de {marcador.nome} ({nome_def}) no duelo 1v1 frente a {portador.nome}!"
                    })

            elif acao == "REMATAR":
                self._resolver_remate(minuto, portador, marcador, gr_def, zona, st_atk, st_def, nome_atk, nome_def)
                posse = chave_def
                zona = "Defesa"

        # Fim do jogo: calcula percentagens e tabela de estatísticas avançadas
        total_ticks = max(1, self.stats["A"]["ticks_posse"] + self.stats["B"]["ticks_posse"])
        for k in ("A", "B"):
            s = self.stats[k]
            s["posse_pct"] = round((s["ticks_posse"] / total_ticks) * 100.0, 1)
            s["passes_pct"] = round((s["passes_completos"] / max(1, s["passes_tentados"])) * 100.0, 1)

        self.narracao.append({
            "minuto": 60,
            "tipo": "FIM",
            "texto": f"Apito final! {self.nome_a} {self.stats['A']['golos']} - {self.stats['B']['golos']} {self.nome_b}!"
        })

        return {
            "narracao": self.narracao,
            "estatisticas": self.stats,
            "tabela_formatada": self._gerar_tabela_consola()
        }

    def _resolver_remate(self, minuto, atirador, defensor, gr_adv, contexto_zona, st_atk, st_def, nome_atk, nome_def, assistente=None):
        """Resolve a matemática de um remate tendo em conta a zona do campo, pressão defensiva
        e reflexos/posicionamento do Guarda-Redes."""
        st_atk["remates_total"] += 1
        atirador.aplicar_desgaste(0.70)

        # Modificador de dificuldade consoante a zona
        if contexto_zona == "Meio-Campo":
            mod_zona = 0.78  # Remate de longe é mais difícil
        elif contexto_zona == "Cara-a-Cara":
            mod_zona = 1.14  # Isolado após drible ou rutura
        else:
            mod_zona = 0.96  # Remate normal dentro da zona de ataque

        # Fórmula Ponderada de Qualidade do Remate:
        # (Remate * 0.65 + Drible * 0.20 + Físico * 0.15) * mod_zona
        mu_remate = (atirador.remate * 0.65 + atirador.drible * 0.20 + atirador.fisico * 0.15) * mod_zona
        val_remate = self._gauss_roll(mu_remate, 9.5)

        # 1. Verifica se o remate vai enquadrado com a baliza (precisão sob pressão do defesa)
        pressao_def = (defensor.defesa * 0.5 + defensor.fisico * 0.5) * 0.22
        limiar_baliza = 46.0 + pressao_def

        if val_remate < limiar_baliza:
            self.narracao.append({
                "minuto": minuto, "tipo": "REMATE_FORA",
                "texto": f"[{minuto}'] {atirador.nome} ({nome_atk}) dispara forte, mas a bola sai ao lado da baliza defendida por {gr_adv.nome}!"
            })
            return

        # O remate vai à baliza! Confronto direto contra o Guarda-Redes
        st_atk["remates_baliza"] += 1
        val_gr = self._gauss_roll(gr_adv.habilidade_gr * 0.98, 8.5)

        if val_remate > val_gr + 2.5:
            st_atk["golos"] += 1
            detalhe_assist = f" (assistência de {assistente.nome})" if assistente else ""
            self.narracao.append({
                "minuto": minuto, "tipo": "GOLO",
                "texto": f"[{minuto}'] GOOOLO de {nome_atk}! {atirador.nome} finaliza sem hipóteses para {gr_adv.nome}{detalhe_assist}! ({self.stats['A']['golos']}-{self.stats['B']['golos']})"
            })
        else:
            st_def["defesas_gr"] += 1
            self.narracao.append({
                "minuto": minuto, "tipo": "DEFESA_GR",
                "texto": f"[{minuto}'] ENORME DEFESA de {gr_adv.nome} ({nome_def})! Voa para negar o golo a {atirador.nome}!"
            })

    def _gerar_tabela_consola(self):
        """Gera a tabela formatada detalhada para impressão na consola e visualização no frontend."""
        sa = self.stats["A"]
        sb = self.stats["B"]
        rem_a = f"{sa['remates_total']} ({sa['remates_baliza']})"
        rem_b = f"{sb['remates_total']} ({sb['remates_baliza']})"
        pas_a = f"{sa['passes_completos']}/{sa['passes_tentados']} ({sa['passes_pct']}%)"
        pas_b = f"{sb['passes_completos']}/{sb['passes_tentados']} ({sb['passes_pct']}%)"
        pos_a = f"{sa['posse_pct']}%"
        pos_b = f"{sb['posse_pct']}%"

        linhas = [
            "=" * 74,
            f"{'ESTATÍSTICAS AVANÇADAS DO JOGO':^74}",
            "=" * 74,
            f"{sa['nome'][:24]:<25} | {'MÉTRICA':^20} | {sb['nome'][:24]:>23}",
            "-" * 74,
            f"{sa['golos']:<25} | {'Golos':^20} | {sb['golos']:>23}",
            f"{pos_a:<25} | {'Posse de Bola (%)':^20} | {pos_b:>23}",
            f"{rem_a:<25} | {'Remates (Baliza)':^20} | {rem_b:>23}",
            f"{sa['defesas_gr']:<25} | {'Defesas do GR':^20} | {sb['defesas_gr']:>23}",
            f"{pas_a:<25} | {'Passes (Comp/Tent %)':^20} | {pas_b:>23}",
            f"{sa['desarmes']:<25} | {'Desarmes':^20} | {sb['desarmes']:>23}",
            f"{sa['intercecoes']:<25} | {'Interceções':^20} | {sb['intercecoes']:>23}",
            "=" * 74
        ]
        tabela_str = "\n".join(linhas)
        print(tabela_str)
        return tabela_str


def match_engine_online(nome_eq_a, alinhamento_a, tatica_a, nome_eq_b, alinhamento_b, tatica_b):
    """Função principal invocada para correr a simulação detalhada por eventos do Modo Online."""
    motor = MatchEngine(nome_eq_a, alinhamento_a, tatica_a, nome_eq_b, alinhamento_b, tatica_b)
    return motor.simular_partida()


# ==============================================================================
# 6. ROTAS HTTP E API (MODO OFFLINE + HEALTHCHECK PARA LINK SEMPRE ABERTO)
# ==============================================================================
@app.route("/")
def index():
    sala_convite = request.args.get("sala", "")
    return render_template("index.html", sala_convite=sala_convite)


@app.route("/ping")
def ping():
    """Endpoint leve para o UptimeRobot / Cron-Job manter o servidor público sempre acordado 24/7."""
    return jsonify({"status": "ok", "jogo": "AC Travanca", "online": True})


@app.route("/api/offline/sortear", methods=["POST"])
def api_offline_sortear():
    """Sorteia 5 jogadores disponíveis (que ainda não foram escolhidos pelo utilizador).
    Um jogador sorteado mas não escolhido pode voltar a aparecer; um escolhido nunca mais aparece."""
    dados = request.get_json() or {}
    escolhidos_ids = set(int(x) for x in dados.get("escolhidos_ids", []))

    todos = carregar_jogadores()
    disponiveis = [j for j in todos if j["id"] not in escolhidos_ids]
    amostra = random.sample(disponiveis, k=min(5, len(disponiveis)))

    return jsonify({
        "opcoes": [sanitizar_jogador_publico(j) for j in amostra],
        "total_disponiveis": len(disponiveis)
    })


@app.route("/api/offline/simular", methods=["POST"])
def api_offline_simular():
    """Recebe os 7 jogadores distribuídos nas posições 1 a 7, a tática escolhida (1, 2 ou 3)
    e o nome do utilizador, corre a Liga Offline e devolve a classificação, melhores marcadores
    e o Top 10 Global de todos os utilizadores que já jogaram."""
    dados = request.get_json() or {}
    alinhamento = dados.get("alinhamento", {})
    tatica = int(dados.get("tatica", 1))
    nome_utilizador = dados.get("nome_utilizador", "Treinador")

    if len(alinhamento) != 7:
        return jsonify({"erro": "Tens de posicionar exatamente os 7 jogadores no campo."}), 400

    resultado = simular_liga_offline(alinhamento, tatica, nome_utilizador)
    return jsonify(resultado)


@app.route("/api/offline/top10", methods=["GET"])
def api_offline_top10():
    """Retorna o Top 10 Global Histórico de todos os utilizadores."""
    return jsonify(obter_top_10_global())


# ==============================================================================
# 7. SALAS PRIVADAS E WEBSOCKETS (MODO ONLINE COM FLASK-SOCKETIO)
# ==============================================================================
SALAS_ONLINE = {}


def sortear_5_sala(sala):
    todos = carregar_jogadores()
    indisponiveis = set(sala["escolhidos_global"])
    disponiveis = [j for j in todos if j["id"] not in indisponiveis]
    amostra = random.sample(disponiveis, k=min(5, len(disponiveis)))
    return [sanitizar_jogador_publico(j) for j in amostra]


def serializar_estado_sala(sala):
    return {
        "room_id": sala["room_id"],
        "estado": sala["estado"],
        "turno_atual": sala["turno_atual"],
        "opcoes_atuais": sala["opcoes_atuais"],
        "jogadores": {
            "P1": {
                "nome": sala["P1"]["nome"],
                "conectado": sala["P1"]["sid"] is not None,
                "equipa": sala["P1"]["equipa"],
                "pronto_tatica": sala["P1"]["pronto_tatica"]
            },
            "P2": {
                "nome": sala["P2"]["nome"],
                "conectado": sala["P2"]["sid"] is not None,
                "equipa": sala["P2"]["equipa"],
                "pronto_tatica": sala["P2"]["pronto_tatica"]
            }
        }
    }


@app.route("/api/online/criar_sala", methods=["POST"])
def api_criar_sala():
    """Gera um link de convite único através de UUID para criar uma sala privada."""
    room_id = str(uuid.uuid4())[:8]
    SALAS_ONLINE[room_id] = {
        "room_id": room_id,
        "estado": "AGUARDANDO_JOGADORES",
        "turno_atual": "P1",
        "escolhidos_global": [],
        "opcoes_atuais": [],
        "P1": {"sid": None, "nome": "Jogador 1", "equipa": [], "tatica": 1, "alinhamento": {}, "pronto_tatica": False},
        "P2": {"sid": None, "nome": "Jogador 2", "equipa": [], "tatica": 1, "alinhamento": {}, "pronto_tatica": False}
    }
    return jsonify({"room_id": room_id})


@socketio.on("entrar_sala")
def handle_entrar_sala(data):
    room_id = str(data.get("room_id", "")).strip()
    nome = str(data.get("nome", "Treinador")).strip()[:30] or "Treinador"

    if room_id not in SALAS_ONLINE:
        SALAS_ONLINE[room_id] = {
            "room_id": room_id,
            "estado": "AGUARDANDO_JOGADORES",
            "turno_atual": "P1",
            "escolhidos_global": [],
            "opcoes_atuais": [],
            "P1": {"sid": None, "nome": "Jogador 1", "equipa": [], "tatica": 1, "alinhamento": {}, "pronto_tatica": False},
            "P2": {"sid": None, "nome": "Jogador 2", "equipa": [], "tatica": 1, "alinhamento": {}, "pronto_tatica": False}
        }

    sala = SALAS_ONLINE[room_id]
    join_room(room_id)

    papel = None
    if sala["P1"]["sid"] is None or sala["P1"]["sid"] == request.sid:
        sala["P1"]["sid"] = request.sid
        sala["P1"]["nome"] = nome
        papel = "P1"
    elif sala["P2"]["sid"] is None or sala["P2"]["sid"] == request.sid:
        sala["P2"]["sid"] = request.sid
        sala["P2"]["nome"] = nome
        papel = "P2"
    else:
        emit("erro_sala", {"mensagem": "Esta sala já está cheia (2 jogadores)."})
        return

    emit("papel_atribuido", {"papel": papel, "room_id": room_id})

    # Se ambos os jogadores estão ligados e o draft ainda não começou, inicia o Draft!
    if sala["P1"]["sid"] and sala["P2"]["sid"] and sala["estado"] == "AGUARDANDO_JOGADORES":
        sala["estado"] = "EM_DRAFT"
        sala["turno_atual"] = "P1"
        sala["opcoes_atuais"] = sortear_5_sala(sala)

    socketio.emit("atualizar_sala", serializar_estado_sala(sala), room=room_id)


@socketio.on("escolher_jogador_draft")
def handle_escolher_draft(data):
    room_id = str(data.get("room_id", ""))
    papel = data.get("papel")
    jogador_id = int(data.get("jogador_id", 0))

    if room_id not in SALAS_ONLINE:
        return
    sala = SALAS_ONLINE[room_id]

    if sala["estado"] != "EM_DRAFT" or sala["turno_atual"] != papel:
        return

    opcoes_ids = [j["id"] for j in sala["opcoes_atuais"]]
    if jogador_id not in opcoes_ids or jogador_id in sala["escolhidos_global"]:
        return

    jog_obj = next(j for j in sala["opcoes_atuais"] if j["id"] == jogador_id)
    sala[papel]["equipa"].append(jog_obj)
    sala["escolhidos_global"].append(jogador_id)

    # Verifica se ambos já têm 7 jogadores (14 escolhas ao todo)
    if len(sala["P1"]["equipa"]) == 7 and len(sala["P2"]["equipa"]) == 7:
        sala["estado"] = "ESCOLHA_TATICA"
        sala["opcoes_atuais"] = []
    else:
        sala["turno_atual"] = "P2" if papel == "P1" else "P1"
        sala["opcoes_atuais"] = sortear_5_sala(sala)

    socketio.emit("atualizar_sala", serializar_estado_sala(sala), room=room_id)


@socketio.on("submeter_tatica_online")
def handle_submeter_tatica(data):
    room_id = str(data.get("room_id", ""))
    papel = data.get("papel")
    tatica = int(data.get("tatica", 1))
    alinhamento = data.get("alinhamento", {})

    if room_id not in SALAS_ONLINE:
        return
    sala = SALAS_ONLINE[room_id]

    if len(alinhamento) != 7:
        return

    sala[papel]["tatica"] = tatica
    sala[papel]["alinhamento"] = alinhamento
    sala[papel]["pronto_tatica"] = True

    socketio.emit("atualizar_sala", serializar_estado_sala(sala), room=room_id)

    # Quando ambos confirmam a formação tática, arranca a simulação online!
    if sala["P1"]["pronto_tatica"] and sala["P2"]["pronto_tatica"]:
        sala["estado"] = "SIMULACAO_CONCLUIDA"
        resultado_jogo = match_engine_online(
            sala["P1"]["nome"], sala["P1"]["alinhamento"], sala["P1"]["tatica"],
            sala["P2"]["nome"], sala["P2"]["alinhamento"], sala["P2"]["tatica"]
        )
        socketio.emit("resultado_jogo_online", resultado_jogo, room=room_id)


init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    socketio.run(app, host="0.0.0.0", port=port, debug=False, allow_unsafe_werkzeug=True)

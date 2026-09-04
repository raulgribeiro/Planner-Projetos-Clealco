# -*- coding: utf-8 -*-
"""
update_dashboard.py
--------------------
Le os arquivos Excel exportados do Planner (Agrícola, Manutenção
Agrícola, Indústria/Clementina+Queiroz, Manutenção Industrial),
extrai a estrutura Projeto -> Fase (DMAIC/PDCA) -> Tarefas, atualiza
os dados embutidos no index.html do dashboard, e sobe (commit + push)
para o repositorio no GitHub.

>>> AJUSTE OS CAMINHOS NA SECAO "CONFIGURACAO" ABAIXO ANTES DE USAR <<<

Requisitos (ja vem prontos no seu ambiente Python normal):
    pip install pandas openpyxl
"""

import json
import re
import subprocess
import sys
from pathlib import Path

import pandas as pd

# ============================================================
# CONFIGURACAO — edite estes 4 caminhos para o seu computador
# ============================================================

# Pasta onde voce salva os Excels exportados do Planner
PASTA_EXCELS = Path(r"C:\Users\raulribeiro\Documents\GitHub\App Planner")

# Nome dos arquivos dentro da pasta acima (exatamente como voce salva)
ARQUIVO_CLEMENTINA = "PROJETOS - PDCA - IND.CLE.xlsx"
ARQUIVO_QUEIROZ = "PROJETOS - PDCA - IND.QRZ.xlsx"
ARQUIVO_MAN_INDUSTRIAL = "PROJETOS - PDCA - MAN.IND.xlsx"
ARQUIVO_MAN_AGRICOLA = "PROJETOS - PDCA - MAN.AGR.xlsx"
ARQUIVO_AGRICOLA = "PROJETOS - PDCA - AGRÍCOLA.xlsx"

# Pasta onde fica o clone local do repositorio GitHub do dashboard.
# No seu caso e a MESMA pasta dos Excels (o index.html fica junto).
PASTA_REPO = PASTA_EXCELS

# ============================================================
# Nao precisa mexer daqui pra baixo
# ============================================================

ARQUIVO_INDEX = PASTA_REPO / "index.html"


def dur_para_int(s):
    if pd.isna(s):
        return None
    m = re.match(r"([\d.,]+)\s*dias?", str(s))
    return int(float(m.group(1).replace(",", "."))) if m else None


def data_str(v):
    if pd.isna(v):
        return None
    return pd.Timestamp(v).strftime("%Y-%m-%d")


# Nomes de fase reconhecidos (como aparecem em CAIXA ALTA na planilha do
# Planner) -> nome canonico usado no dashboard. Inclui variantes/typos
# ja vistos nos exports (ex.: "EXECULTAR" na planilha da Agrícola).
FASES_CANONICAS = {
    "DEFINIR": "DEFINIR", "MEDIR": "MEDIR", "ANALISAR": "ANALISAR",
    "IMPLEMENTAR": "IMPLEMENTAR", "CONTROLAR": "CONTROLAR",
    "CONTRATO": "CONTRATO", "PLANEJAR": "PLANEJAR",
    "EXECUTAR": "EXECUTAR", "EXECULTAR": "EXECUTAR",
    "AGIR": "AGIR",
}


def _nome_compacto(nome):
    return re.sub(r"[^A-Z0-9]", "", nome.upper())


def _eh_categoria(nome):
    """Linha que agrupa varios projetos por metodologia (ex.: uma linha
    'GREEN BELT' ou 'LEAN MANUFACTURE' sozinha, sem nome de responsavel
    junto — vista na planilha da Agrícola)."""
    return _nome_compacto(nome) in ("GREENBELT", "LEANMANUFACTURE", "LEANMANUFACTURING")


def _eh_fase(nome):
    """So conta como fase se o nome ja vier em CAIXA ALTA na planilha
    (padrao real de DEFINIR/MEDIR/...). Uma tarefa avulsa chamada
    'Contrato' (Title Case, comum dentro de DEFINIR/CONTROLAR) nao passa
    nesse teste, evitando confundi-la com a fase CONTRATO do Lean."""
    limpo = nome.strip()
    return limpo == limpo.upper() and limpo.upper() in FASES_CANONICAS


def _nome_canonico_fase(nome):
    return FASES_CANONICAS.get(nome.strip().upper(), nome.strip())


def _tipo_por_nome(nome):
    nc = _nome_compacto(nome)
    if "GREENBELT" in nc:
        return "GREENBELT"
    if "LEAN" in nome.upper():
        return "LEAN MANUFACTURING"
    return "OUTRO"


def extrair_projetos(caminho_xlsx):
    """Le um export do Planner e devolve a lista de projetos com suas fases.

    Robusto a duas variacoes de layout vistas nos exports:
    1. Padrao (Clementina/Queiroz/Manutencoes): projeto no nivel raiz
       (ex. '1'), com um subgrupo intermediario de um so filho (RESIDUARIA/
       NIR/CONTRATO) antes das fases DEFINIR/MEDIR/...
    2. Agrícola: um nivel extra de categoria por metodologia ('GREEN BELT'
       / 'LEAN MANUFACTURE'), com varios projetos reais como filhos diretos
       dela, e as vezes fases aninhadas de forma torta (ex. IMPLEMENTAR e
       CONTROLAR dentro de ANALISAR, ou CONTROLAR dentro de IMPLEMENTAR).
    """
    raw = pd.read_excel(caminho_xlsx, sheet_name=0, header=None)

    header_row_idx = None
    for i in range(raw.shape[0]):
        if str(raw.iloc[i, 0]).strip() == "Número de tarefa":
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError(
            f"Nao encontrei a linha de cabecalho ('Número de tarefa') em {caminho_xlsx}. "
            "O layout do export do Planner pode ter mudado."
        )

    headers = raw.iloc[header_row_idx].tolist()
    df = raw.iloc[header_row_idx + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    nodes = {}
    children = {}
    for _, row in df.iterrows():
        nivel = str(row["Número do nível hierárquico"]).strip()
        if nivel in ("nan", "None", ""):
            continue
        path = tuple(nivel.split("."))
        nome = str(row["Nome"]).strip()
        inicio = data_str(row["Início"])
        fim = data_str(row["Concluir"])
        duracao = dur_para_int(row["Duração"])
        perc = row["% concluída"]
        perc = None if pd.isna(perc) else round(float(perc) * 100, 1)
        nodes[path] = {"nome": nome, "inicio": inicio, "fim": fim, "duracao": duracao, "perc": perc}
        children.setdefault(path[:-1], []).append(path)

    def coletar_fases(candidatos):
        resultado = []
        for cpath in candidatos:
            cnome = nodes[cpath]["nome"]
            if _eh_fase(cnome):
                resultado.append(cpath)
                resultado.extend(coletar_fases(children.get(cpath, [])))
        return resultado

    def montar_projeto(path, tipo_contexto):
        node = nodes[path]
        candidatos = children.get(path, [])
        # subgrupo intermediario de um so filho (RESIDUARIA/NIR/CONTRATO)
        if len(candidatos) == 1 and not _eh_fase(nodes[candidatos[0]]["nome"]):
            candidatos = children.get(candidatos[0], [])
        fases_paths = coletar_fases(candidatos)

        fases = []
        for fp in fases_paths:
            fnode = nodes[fp]
            filhos_fase = children.get(fp, [])
            tarefas = [c for c in filhos_fase if not _eh_fase(nodes[c]["nome"])]
            concluidas = sum(1 for c in tarefas if nodes[c]["perc"] is not None and nodes[c]["perc"] >= 100)
            fases.append({
                "nome": _nome_canonico_fase(fnode["nome"]),
                "inicio": fnode["inicio"], "fim": fnode["fim"], "duracao": fnode["duracao"],
                "percReal": fnode["perc"], "tarefasTotal": len(tarefas), "tarefasConcluidas": concluidas,
            })

        nome_projeto = node["nome"]
        responsavel = nome_projeto.split("-", 1)[0].strip()
        return {
            "nome": nome_projeto, "tipo": tipo_contexto, "responsavel": responsavel,
            "inicio": node["inicio"], "fim": node["fim"], "duracao": node["duracao"],
            "percReal": node["perc"], "fases": fases,
        }

    raizes = children.get((), [])
    projetos = []
    for raiz_path in raizes:
        raiz = nodes[raiz_path]
        if _eh_categoria(raiz["nome"]):
            tipo_contexto = _tipo_por_nome(raiz["nome"])
            for filho_path in children.get(raiz_path, []):
                projetos.append(montar_projeto(filho_path, tipo_contexto))
        else:
            projetos.append(montar_projeto(raiz_path, _tipo_por_nome(raiz["nome"])))

    return projetos


def substituir_bloco(html, marcador_inicio, marcador_fim, nova_variavel, dados):
    padrao = re.compile(
        re.escape(marcador_inicio) + r".*?" + re.escape(marcador_fim),
        re.DOTALL,
    )
    novo_bloco = (
        f"{marcador_inicio}\n"
        f"const {nova_variavel} = {json.dumps(dados, ensure_ascii=False)};\n"
        f"{marcador_fim}"
    )
    novo_html, n = padrao.subn(novo_bloco, html)
    if n == 0:
        raise ValueError(
            f"Nao encontrei os marcadores {marcador_inicio} / {marcador_fim} no index.html. "
            "Confira se o arquivo foi editado manualmente."
        )
    return novo_html


def rodar_git(comando, cwd):
    resultado = subprocess.run(
        comando, cwd=cwd, capture_output=True, text=True, shell=False
    )
    print(">", " ".join(comando))
    if resultado.stdout.strip():
        print(resultado.stdout.strip())
    if resultado.returncode != 0:
        print(resultado.stderr.strip(), file=sys.stderr)
    return resultado.returncode == 0


def main():
    caminho_clementina = PASTA_EXCELS / ARQUIVO_CLEMENTINA
    caminho_queiroz = PASTA_EXCELS / ARQUIVO_QUEIROZ
    caminho_man_ind = PASTA_EXCELS / ARQUIVO_MAN_INDUSTRIAL
    caminho_man_agr = PASTA_EXCELS / ARQUIVO_MAN_AGRICOLA
    caminho_agricola = PASTA_EXCELS / ARQUIVO_AGRICOLA

    def ler(caminho, rotulo):
        if not caminho.exists():
            print(f"[AVISO] Nao achei {caminho} — pulando {rotulo}.")
            return None
        print(f"Lendo {caminho} ...")
        dados = extrair_projetos(caminho)
        print(f"  {len(dados)} projeto(s) encontrados.")
        return dados

    dados_clementina = ler(caminho_clementina, "Clementina")
    dados_queiroz = ler(caminho_queiroz, "Queiroz")
    dados_man_ind = ler(caminho_man_ind, "Manutenção Industrial")
    dados_man_agr = ler(caminho_man_agr, "Manutenção Agrícola")
    dados_agricola = ler(caminho_agricola, "Agrícola")

    if all(d is None for d in (dados_clementina, dados_queiroz, dados_man_ind, dados_man_agr, dados_agricola)):
        print("Nenhum dos Excels foi encontrado. Nada para atualizar. Encerrando.")
        if PASTA_EXCELS.exists():
            arquivos = sorted(p.name for p in PASTA_EXCELS.iterdir() if p.suffix.lower() == ".xlsx")
            if arquivos:
                print(f"\nArquivos .xlsx que EXISTEM em {PASTA_EXCELS}:")
                for nome in arquivos:
                    print(f"  - {nome}")
                print("\nCompare com os nomes esperados no topo do script (ARQUIVO_CLEMENTINA, ARQUIVO_QUEIROZ, etc.)")
            else:
                print(f"\nA pasta {PASTA_EXCELS} existe, mas nao tem nenhum arquivo .xlsx dentro.")
        else:
            print(f"\n[ERRO] A pasta {PASTA_EXCELS} nem existe. Confira o caminho em PASTA_EXCELS no topo do script.")
        return

    if not ARQUIVO_INDEX.exists():
        print(f"[ERRO] index.html nao encontrado em {ARQUIVO_INDEX}.")
        print("Confira o caminho PASTA_REPO no topo deste script.")
        return

    html = ARQUIVO_INDEX.read_text(encoding="utf-8")

    if dados_clementina is not None:
        html = substituir_bloco(
            html,
            "// ===DADOS_CLEMENTINA_INICIO===",
            "// ===DADOS_CLEMENTINA_FIM===",
            "projetosClementina",
            dados_clementina,
        )
    if dados_queiroz is not None:
        html = substituir_bloco(
            html,
            "// ===DADOS_QUEIROZ_INICIO===",
            "// ===DADOS_QUEIROZ_FIM===",
            "projetosQueiroz",
            dados_queiroz,
        )
    if dados_man_ind is not None:
        html = substituir_bloco(
            html,
            "// ===DADOS_MANUTENCAO_INDUSTRIAL_INICIO===",
            "// ===DADOS_MANUTENCAO_INDUSTRIAL_FIM===",
            "projetosManutencaoIndustrial",
            dados_man_ind,
        )
    if dados_man_agr is not None:
        html = substituir_bloco(
            html,
            "// ===DADOS_MANUTENCAO_AGRICOLA_INICIO===",
            "// ===DADOS_MANUTENCAO_AGRICOLA_FIM===",
            "projetosManutencaoAgricola",
            dados_man_agr,
        )
    if dados_agricola is not None:
        html = substituir_bloco(
            html,
            "// ===DADOS_AGRICOLA_INICIO===",
            "// ===DADOS_AGRICOLA_FIM===",
            "projetosAgricola",
            dados_agricola,
        )

    ARQUIVO_INDEX.write_text(html, encoding="utf-8")
    print(f"index.html atualizado em {ARQUIVO_INDEX}")

    # Publica no GitHub
    if not rodar_git(["git", "add", "index.html"], cwd=PASTA_REPO):
        print("[ERRO] Falhou o 'git add'. Confira se PASTA_REPO e um clone git valido.")
        return

    from datetime import datetime
    mensagem = f"Atualiza dados dos projetos PDCA — {datetime.now().strftime('%d/%m/%Y %H:%M')}"

    commit_ok = rodar_git(["git", "commit", "-m", mensagem], cwd=PASTA_REPO)
    if not commit_ok:
        print("Nada novo para commitar (dados iguais aos ja publicados) ou houve erro acima.")
        return

    if rodar_git(["git", "push"], cwd=PASTA_REPO):
        print("Publicado com sucesso! O GitHub Pages atualiza em 1-2 minutos.")
    else:
        print("[ERRO] Falhou o 'git push'. Confira sua autenticacao do git/GitHub.")


if __name__ == "__main__":
    main()

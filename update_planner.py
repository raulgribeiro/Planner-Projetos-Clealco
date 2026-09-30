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
import os
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

# No GitHub Actions o link do JSON exportado pelo Power Automate vem por
# variavel de ambiente; nesse modo o script roda dentro do proprio repo e
# NAO faz git (o workflow faz o commit/push).
PLANNER_SHAREPOINT_URL = os.environ.get("PLANNER_SHAREPOINT_URL")
if PLANNER_SHAREPOINT_URL:
    PASTA_REPO = Path(__file__).parent.resolve()

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

    return _projetos_de_nos(nodes, children)


# Nome do plano no Planner Premium (Dataverse) -> chave usada no main()
PLANOS_DATAVERSE = {
    "PROJETOS | PDCA | IND.CLE": "clementina",
    "PROJETOS | PDCA | IND.QRZ": "queiroz",
    "PROJETOS | PDCA | MAN.IND": "man_ind",
    "PROJETOS | PDCA | MAN.AGR": "man_agr",
    "PROJETOS | PDCA | AGRÍCOLA": "agricola",
    "PROJETOS | PDCA | AGRÍCOLA - SEGUNDO SEMESTRE": "agricola",  # somado ao plano Agrícola
}

# Planos SDCA: plano "PROJETOS | SDCA | <Processo>" = 1 processo; as tarefas de
# nivel 1 sao as etapas. Chave = processo (maiusculo) -> bloco do index.html.
PLANOS_SDCA = {
    "ALMOXARIFADO": "sdca_apoio",
}


def _data_iso(v):
    return v[:10] if isinstance(v, str) and len(v) >= 10 else None


def extrair_projetos_json(rows):
    """Recebe as linhas da tabela Project Tasks (Dataverse) de UM plano e
    devolve a mesma estrutura de extrair_projetos(). O numero hierarquico
    (1, 1.1, 1.1.1 ...) e reconstruido a partir da ordem de exibicao +
    nivel de indentacao, igual ao export do Planner."""
    rows = sorted(rows, key=lambda r: r.get("msdyn_displaysequence") or 0)
    nodes, children, contadores = {}, {}, []
    for r in rows:
        nivel = int(r.get("msdyn_outlinelevel") or 1)
        contadores = contadores[:nivel]
        while len(contadores) < nivel:
            contadores.append(0)
        contadores[nivel - 1] += 1
        path = tuple(str(c) for c in contadores)
        dur = r.get("msdyn_duration")
        prog = r.get("msdyn_progress")
        nodes[path] = {
            "nome": str(r.get("msdyn_subject") or "").strip(),
            "inicio": _data_iso(r.get("msdyn_start")),
            "fim": _data_iso(r.get("msdyn_finish")),
            "duracao": None if dur is None else int(round(float(dur))),
            "perc": None if prog is None else round(float(prog) * 100, 1),
        }
        children.setdefault(path[:-1], []).append(path)
    return _projetos_de_nos(nodes, children)


def projeto_sdca(nome_processo, rows):
    """Plano SDCA -> lista com 1 projeto (o processo) e suas etapas (tarefas
    de nivel 1). % do processo = media das etapas ponderada pela duracao."""
    from datetime import date
    rows = sorted(rows, key=lambda r: r.get("msdyn_displaysequence") or 0)
    etapas = []
    for r in rows:
        if int(r.get("msdyn_outlinelevel") or 1) != 1:
            continue
        dur = r.get("msdyn_duration")
        prog = r.get("msdyn_progress")
        perc = None if prog is None else round(float(prog) * 100, 1)
        etapas.append({
            "nome": str(r.get("msdyn_subject") or "").strip(),
            "inicio": _data_iso(r.get("msdyn_start")), "fim": _data_iso(r.get("msdyn_finish")),
            "duracao": None if dur is None else int(round(float(dur))),
            "percReal": perc, "tarefasTotal": 1,
            "tarefasConcluidas": 1 if perc is not None and perc >= 100 else 0,
        })
    if not etapas:
        return []
    inicios = [e["inicio"] for e in etapas if e["inicio"]]
    fins = [e["fim"] for e in etapas if e["fim"]]
    ini, fim = min(inicios), max(fins)
    dias = (date.fromisoformat(fim) - date.fromisoformat(ini)).days + 1
    pesos = [(e["duracao"] or 0) for e in etapas]
    soma = sum(pesos)
    perc = sum((e["percReal"] or 0) * w for e, w in zip(etapas, pesos)) / soma if soma else 0
    return [{
        "nome": nome_processo, "tipo": "SDCA", "responsavel": "",
        "inicio": ini, "fim": fim, "duracao": dias,
        "percReal": float(round(perc)), "fases": etapas,
    }]


def baixar_planner_json(url):
    """Baixa o JSON exportado pelo Power Automate (link 'Qualquer pessoa' do
    SharePoint/OneDrive, sem login): 2 requisicoes na mesma sessao."""
    import requests
    s = requests.Session()
    s.headers["User-Agent"] = "Mozilla/5.0"
    s.get(url, timeout=60)
    sep = "&" if "?" in url else "?"
    resp = s.get(f"{url}{sep}download=1", timeout=180)
    resp.raise_for_status()
    try:
        return resp.json()["value"]
    except Exception as e:
        raise RuntimeError(
            "O download do SharePoint nao retornou o JSON do Planner. O link "
            f"pode ter expirado/mudado de permissao ({e})."
        )


def projetos_por_plano(linhas):
    """Agrupa as linhas do Dataverse por plano e devolve {chave: projetos}."""
    por_plano = {}
    for r in linhas:
        nome = r.get("_msdyn_project_value@OData.Community.Display.V1.FormattedValue", "")
        por_plano.setdefault(nome.strip(), []).append(r)
    saida = {}
    for nome, rows in por_plano.items():
        if "SDCA" in nome.upper():
            processo = nome.split("|")[-1].strip()
            chave_sdca = PLANOS_SDCA.get(processo.upper())
            if chave_sdca:
                saida.setdefault(chave_sdca, []).extend(projeto_sdca(processo, rows))
                print(f"  {nome}: {len(rows)} tarefas -> processo SDCA {processo!r}")
                continue
        chave = PLANOS_DATAVERSE.get(nome)
        if chave is None:
            print(f"[AVISO] Plano desconhecido ignorado: {nome!r} ({len(rows)} tarefas)")
            continue
        novos = extrair_projetos_json(rows)
        saida.setdefault(chave, []).extend(novos)
        print(f"  {nome}: {len(rows)} tarefas -> {len(novos)} projeto(s)")
    return saida


def _projetos_de_nos(nodes, children):
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
    novo_html, n = padrao.subn(lambda _m: novo_bloco, html)  # lambda: preserva barras invertidas do JSON
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

    if PLANNER_SHAREPOINT_URL:
        print("Baixando tarefas do Planner (JSON do Power Automate)...")
        por_plano = projetos_por_plano(baixar_planner_json(PLANNER_SHAREPOINT_URL))
        # ignora projetos-modelo que ficam dentro dos planos
        for chave, lista in por_plano.items():
            por_plano[chave] = [p for p in lista if not p["nome"].upper().startswith("TEMPLATE")]
        dados_clementina = por_plano.get("clementina")
        dados_queiroz = por_plano.get("queiroz")
        dados_man_ind = por_plano.get("man_ind")
        dados_man_agr = por_plano.get("man_agr")
        dados_agricola = por_plano.get("agricola")
        dados_sdca_apoio = por_plano.get("sdca_apoio")
    else:
        dados_sdca_apoio = None
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

    if dados_sdca_apoio:
        html = substituir_bloco(
            html,
            "// ===DADOS_SDCA_APOIO_INICIO===",
            "// ===DADOS_SDCA_APOIO_FIM===",
            "projetosSdcaApoio",
            dados_sdca_apoio,
        )

    ARQUIVO_INDEX.write_text(html, encoding="utf-8")
    print(f"index.html atualizado em {ARQUIVO_INDEX}")

    if PLANNER_SHAREPOINT_URL:
        print("Modo Actions: commit/push fica por conta do workflow.")
        return

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

"""Simulacao de IPTU.

Uso:
    python simulacaoIPTU.py --tipo territorial  [--entrada X.xlsx] [--saida Y.xlsx] [--ano 2026]
    python simulacaoIPTU.py --tipo completa     ...
    python simulacaoIPTU.py --tipo sem_unidades ...
    python simulacaoIPTU.py --tipo novo_padrao  ...

Tipos de planilha (aba SCRIPT_IPTU_ENTRADA):
  territorial  : so imoveis territoriais. Calcula apenas o terreno (VVT) e usa a
                 faixa NAO EDIFICADO, ignorando qualquer edificacao.
  completa     : todas as unidades do imovel ja estao na planilha (aerolevantamento).
                 A Betha so completa campos vazios de cadastro/terreno e os valores
                 Anterior/Atual; os dados de construcao vem apenas da planilha.
  sem_unidades : a planilha traz so a unidade principal (edificada ou nao). As demais
                 unidades ativas sao buscadas na Betha (idImovelPrincipal) com material,
                 tipo, padrao, uso e area construida. Na principal, o que estiver
                 preenchido na planilha prevalece; o vazio e completado pela Betha.
  novo_padrao  : igual a sem_unidades, mas o Padrao_Digitado da principal (planilha)
                 passa a valer em TODAS as unidades do imovel (o padrao anterior de
                 cada unidade fica registrado em Observacoes).

Em todos os tipos a planilha prevalece sobre a Betha; divergencias de Secao e
Inscricao sao registradas na coluna Observacoes.

Calculo:
  VVT   = Area_Terreno (unidade principal) * m2 da PGV (Distrito/Setor/Zona/Quadra/Secao)
  VVE-U = Area_Construida_Unidade_InLoco * m2 da PGP (material/tipo/padrao)
  VVTE  = soma das VVE-U do imovel ; VV_T = VVT + VVTE
  Aliq  = faixa por Uso_da_Edificacao da principal e VV_T (sem edificacao -> NAO EDIFICADO)
  VI    = VV_T * Aliq / 100 ; Impacto = VI - Valor_Imposto_Atual
"""
import argparse
import re
import sys
import unicodedata

import numpy as np
import pandas as pd

from betha_api import Betha, linha_do_imovel

ABA_ENTRADA = "SCRIPT_IPTU_ENTRADA"
TIPOS = ("territorial", "completa", "sem_unidades", "novo_padrao")
COM_UNIDADES_BETHA = ("sem_unidades", "novo_padrao")  # buscam as subordinadas na Betha

COLUNAS_SAIDA = [
    "Codigo", "Codigo_Principal", "ID", "ID_Principal", "Inscricao_imobiliaria", "Inscricao_imobiliaria_InLoco",
    "Unidade", "Secao", "Tipo_de_Edificacao", "Tipo_de_Material", "Uso_da_Edificacao",
    "Area_Construida_Unidade_InLoco", "Padrao_Digitado", "Nome_Proprietario", "Cpf_Cnpj",
    "Bairro", "Tipo", "Situacao", "Area_Edificada_Total", "Area_Construida_Obras",
    "Area_Terreno", "Taxacao", "Numero_de_Pavimentos", "Classificacao", "Patrimonio",
    "Valor_Venal_Terreno_Anterior", "Valor_Venal_Terreno_Atual", "Valor_Venal_Terreno_Simulacao",
    "Valor_Venal_Edificacao_Anterior", "Valor_Venal_Edificacao_Atual",
    "Valor_Venal_Edificacao_Simulacao", "Valor_Venal_Total_Imovel_Anterior",
    "Valor_Venal_Total_Imovel_Atual", "Valor_Venal_Total_Imovel_Simulacao",
    "Valor_Aliquota_Anterior", "Valor_Aliquota_Atual", "Valor_Aliquota_Simulacao",
    "Tipo_de_Imposto", "Atualizacao_Aerolevantamento", "Ano", "Valor_Imposto_Anterior",
    "Valor_Imposto_Atual", "Valor_Imposto_Simulado", "Valor_Impacto",
    "Qtd_Unidades", "Observacoes",
]

# Campos que a Betha pode completar, por tipo de planilha
CAMPOS_CADASTRO = ["Inscricao_imobiliaria", "Secao", "Codigo", "Unidade", "Bairro", "Tipo",
                   "Situacao", "Nome_Proprietario", "Cpf_Cnpj", "Area_Terreno", "Taxacao",
                   "Tipo_de_Imposto", "Classificacao", "Patrimonio", "Numero_de_Pavimentos",
                   "Area_Edificada_Total", "Atualizacao_Aerolevantamento"]
CAMPOS_VALORES = [f"{b}_{s}" for b in ("Valor_Venal_Terreno", "Valor_Venal_Edificacao",
                                       "Valor_Venal_Total_Imovel", "Valor_Aliquota",
                                       "Valor_Imposto") for s in ("Anterior", "Atual")]
CAMPOS_CONSTRUCAO = ["Tipo_de_Material", "Tipo_de_Edificacao", "Padrao_Digitado",
                     "Area_Construida_Unidade_InLoco"]
PREENCHE = {
    "territorial": CAMPOS_CADASTRO + CAMPOS_VALORES,
    "completa": CAMPOS_CADASTRO + CAMPOS_VALORES + ["Uso_da_Edificacao"],
    "sem_unidades": CAMPOS_CADASTRO + CAMPOS_VALORES + ["Uso_da_Edificacao"] + CAMPOS_CONSTRUCAO,
}
PREENCHE["novo_padrao"] = PREENCHE["sem_unidades"]


# =========================================================
# Utilitarios
# =========================================================
def para_numero(s):
    """Aceita 1234.56, 1234,56 e 1.234,56."""
    def conv(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return np.nan
        if isinstance(v, (int, float)):
            return float(v)
        t = str(v).replace("R$", "").strip()
        if t == "" or t.lower() == "nan":
            return np.nan
        if "," in t:
            t = t.replace(".", "").replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return np.nan
    return s.map(conv)


PREFIXO_NUMERICO = re.compile(r"^\s*\d[\d\s.\-–:/)]*")
CAMPOS_COM_CODIGO = ["Tipo_de_Material", "Padrao_Digitado", "Uso_da_Edificacao", "Tipo_de_Edificacao"]


def sem_prefixo(v):
    """Remove o codigo numerico do inicio: "1 ALVENARIA", "01 - MEDIO", "3.RESIDENCIAL",
    "2) CASA RESIDENCIAL" -> so o texto. Valor so com numero vira vazio."""
    if vazio(v):
        return v
    return PREFIXO_NUMERICO.sub("", str(v)).strip()


def chave(v):
    """Normaliza rotulo para casar planilha/Betha com PGP e aliquota:
    "1 CASA-RESIDENCIAL" -> "CASA RESIDENCIAL", "MÉDIO" -> "MEDIO"."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    t = sem_prefixo(str(v).strip())
    t = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode()
    t = re.sub(r"\s+", " ", t.replace("-", " ")).strip().upper()
    return {"NAN": "", "ESTR. METALICA": "METALICA", "COM. E RES.": "COM. E RESD."}.get(t, t)


def texto_id(s):
    return (s.fillna("").astype(str).str.strip()
            .str.replace(r"\.0$", "", regex=True).replace({"nan": "", "0": ""}))


def vazio(v):
    return v is None or (isinstance(v, float) and np.isnan(v)) or str(v).strip() in ("", "nan")


# =========================================================
# Entrada
# =========================================================
def ler_entrada(caminho):
    if str(caminho).lower().endswith(".csv"):  # CSV do relatorio de dados cadastrais (Betha)
        df = pd.read_csv(caminho, sep=";", dtype=str, encoding="iso-8859-1")
    else:
        abas = pd.ExcelFile(caminho).sheet_names
        aba = ABA_ENTRADA if ABA_ENTRADA in abas else abas[0]  # ex.: aba "CGR" das planilhas do aerolevantamento
        df = pd.read_excel(caminho, sheet_name=aba, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    if "ID_Principal" not in df.columns and "Codigo_Principal" in df.columns:
        df = df.rename(columns={"Codigo_Principal": "ID_Principal"})  # layout anterior
    for c in COLUNAS_SAIDA:
        if c not in df.columns:
            df[c] = np.nan
    df = df.astype(object)
    # Relatorio Betha (CAMPOS_SCRIPT-betha) traz a area cadastral em Area_Construida_Unidade
    if "Area_Construida_Unidade" in df.columns:
        vaz = df["Area_Construida_Unidade_InLoco"].map(vazio)
        df.loc[vaz, "Area_Construida_Unidade_InLoco"] = df.loc[vaz, "Area_Construida_Unidade"]
    for c in CAMPOS_COM_CODIGO:
        df[c] = df[c].map(sem_prefixo)
    df["ID"] = texto_id(df["ID"])
    df["ID_Principal"] = texto_id(df["ID_Principal"])
    df["Origem"] = "PLANILHA"
    df["Campos_Betha"] = ""
    df["Observacoes"] = ""
    return df


# =========================================================
# Complemento pela Betha
# =========================================================
def completar_com_betha(df, tipo, api, ano):
    # 1) Linhas sem ID: resolve pelo Codigo
    sem_id = (df["ID"] == "") & (texto_id(df["Codigo"]) != "")
    if sem_id.any():
        por_cod = {str(i["codigo"]): i for i in api.imoveis("codigo", texto_id(df.loc[sem_id, "Codigo"]))}
        for i in df.index[sem_id]:
            im = por_cod.get(texto_id(pd.Series([df.at[i, "Codigo"]]))[0])
            if im:
                df.at[i, "ID"] = str(im["id"])

    ids = set(df.loc[df["ID"] != "", "ID"])
    cadastro = {str(i["id"]): linha_do_imovel(i) for i in api.imoveis("id", ids)}
    for i in df.index:
        reg = cadastro.get(df.at[i, "ID"])
        if reg and df.at[i, "ID_Principal"] == "":
            df.at[i, "ID_Principal"] = reg["ID_Principal"]
    nao_achados = sorted(ids - set(cadastro))
    if nao_achados:
        print(f"ATENCAO: {len(nao_achados)} ID(s) nao encontrados na Betha: {nao_achados[:10]}")

    # 2) sem_unidades: traz a principal (se a planilha trouxe outra unidade) e as subordinadas
    if tipo in COM_UNIDADES_BETHA:
        principais = {p or i for p, i in zip(df["ID_Principal"], df["ID"]) if (p or i)}
        faltam = principais - set(cadastro)
        for im in api.imoveis("id", faltam) if faltam else []:
            cadastro[str(im["id"])] = linha_do_imovel(im)
        for im in api.unidades_subordinadas(principais):
            cadastro[str(im["id"])] = linha_do_imovel(im)
        presentes = set(df["ID"])
        novas = [dict(reg, Origem="BETHA") for idc, reg in cadastro.items()
                 if idc not in presentes and reg["ID_Principal"] in principais]
        if novas:
            df = pd.concat([df, pd.DataFrame(novas)], ignore_index=True).astype(object)
            df[["Campos_Betha", "Observacoes"]] = df[["Campos_Betha", "Observacoes"]].fillna("")
        print(f"Betha: {len(principais)} principal(is), {len(novas)} unidade(s) adicionada(s)")

    # 3) Campos adicionais e preenchimento dos vazios
    adicionais = api.campos_adicionais([int(x) for x in df["ID"] if x], ano)
    permitidos = PREENCHE[tipo]
    for i in df.index:
        idc = df.at[i, "ID"]
        if not idc:
            continue
        betha = dict(cadastro.get(idc, {}))
        betha.update(adicionais.get(int(idc), {}))
        preenchidos, obs = [], []
        for col in ("Secao", "Inscricao_imobiliaria"):
            a, b = df.at[i, col], betha.get(col)
            if not vazio(a) and not vazio(b) and                     str(a).strip().removesuffix(".0").lstrip("0") != str(b).strip().lstrip("0"):
                obs.append(f"{col} planilha {a} x Betha {b}")
        for col in permitidos:
            if vazio(df.at[i, col]) and not vazio(betha.get(col)):
                df.at[i, col] = betha[col]
                preenchidos.append(col)
        if df.at[i, "Origem"] == "BETHA":
            preenchidos = ["unidade inteira"]
        df.at[i, "Campos_Betha"] = ", ".join(preenchidos)
        df.at[i, "Observacoes"] = "; ".join(obs)
    df["Ano"] = df["Ano"].where(df["Ano"].notna(), ano)
    return df


# =========================================================
# Calculo
# =========================================================
def calcular(df, tipo):
    # Imovel e unidade principal
    df["ChaveImovel"] = np.where(df["ID_Principal"] != "", df["ID_Principal"], df["ID"])
    # Unidade nova (in loco, ainda sem cadastro na Betha): entra no imovel cuja linha
    # cadastrada tem a mesma Inscricao_imobiliaria (o englobado da planilha)
    insc = df["Inscricao_imobiliaria"].fillna("").astype(str).str.strip()
    com_chave = (df["ChaveImovel"] != "") & (insc != "")
    chave_insc = df[com_chave].assign(_i=insc[com_chave]).drop_duplicates("_i").set_index("_i")["ChaveImovel"]
    sem_chave = df["ChaveImovel"] == ""
    df.loc[sem_chave, "ChaveImovel"] = insc[sem_chave].map(chave_insc).fillna("")
    sem_chave = df["ChaveImovel"] == ""
    df.loc[sem_chave, "ChaveImovel"] = "LINHA_" + df.index[sem_chave].astype(str)
    # a principal e sempre uma linha cadastrada (com ID), salvo imovel sem nenhum ID
    candidata = (((df["ID_Principal"] == "") | (df["ID_Principal"] == df["ID"])) & (df["ID"] != "")) \
        | df["ChaveImovel"].str.startswith("LINHA_")
    df["_ord"] = pd.to_numeric(df["Unidade"], errors="coerce")
    ordem = df.assign(_nc=~candidata).sort_values(["ChaveImovel", "_nc", "_ord"])
    df["Principal"] = df.index.isin(ordem.groupby("ChaveImovel").head(1).index)
    pri = df["Principal"]

    # ID_Principal / Codigo_Principal: a principal repete o proprio ID/codigo e a
    # unidade englobada recebe o ID/codigo da principal do imovel
    real = ~df["ChaveImovel"].str.startswith("LINHA_")
    df["ID_Principal"] = np.where(real, df["ChaveImovel"], "")
    cod = texto_id(df["Codigo"])
    cod_por_id = cod[df["ID"] == df["ChaveImovel"]].groupby(df["ChaveImovel"]).first()
    cod_da_principal = cod[pri].groupby(df.loc[pri, "ChaveImovel"]).first()
    df["Codigo_Principal"] = (df["ChaveImovel"].map(cod_por_id)
                              .fillna(df["ChaveImovel"].map(cod_da_principal)).fillna(""))

    # novo_padrao: o padrao da principal (planilha) vale em todas as unidades do imovel
    if tipo == "novo_padrao":
        fonte = pri & (df["Origem"] == "PLANILHA") & ~df["Padrao_Digitado"].map(vazio)
        padrao_imovel = df.loc[fonte].groupby("ChaveImovel")["Padrao_Digitado"].first()
        for i in df.index[~pri]:
            novo = padrao_imovel.get(df.at[i, "ChaveImovel"])
            if novo is None or chave(novo) == chave(df.at[i, "Padrao_Digitado"]):
                continue
            antigo = df.at[i, "Padrao_Digitado"]
            df.at[i, "Padrao_Digitado"] = novo
            msg = f"padrao {'' if vazio(antigo) else antigo} -> {novo} (novo padrao da principal)"
            df.at[i, "Observacoes"] = "; ".join(filter(None, [df.at[i, "Observacoes"], msg]))

    # VVT - PGV
    pgv = pd.read_excel("pgv.xlsx", dtype=str)
    partes = (df["Inscricao_imobiliaria"].where(~df["Inscricao_imobiliaria"].map(vazio),
                                                 df["Inscricao_imobiliaria_InLoco"])
              .fillna("").astype(str).str.strip().str.split(".", expand=True)
              .reindex(columns=range(4)))
    for n, c in enumerate(["Distrito", "Setor", "Zona", "Quadra"]):
        df[c] = partes[n].fillna("").str.strip().str.zfill(2)
        pgv[c] = pgv[c].astype(str).str.strip().str.zfill(2)
    DEPARA_SECAO = {"170203": "61796"}  # R. Marina F. de Alessio (Lei 5110/2023) -> secao antiga na PGV
    df["SecaoPgv"] = (texto_id(df["Secao"]).replace(DEPARA_SECAO).str.zfill(5))
    pgv["SecaoPgv"] = pgv["Secao"].astype(str).str.strip().str.zfill(5)
    k = ["Distrito", "Setor", "Zona", "Quadra", "SecaoPgv"]
    pgv_u = pgv[k + ["Valor"]].drop_duplicates(subset=k).rename(columns={"Valor": "ValorPgv"})
    df = df.merge(pgv_u, on=k, how="left", validate="m:1")
    pri = df["Principal"]

    sem_pgv = pri & df["ValorPgv"].isna()
    if sem_pgv.any():
        val_quadra = (pgv.groupby(["Distrito", "Setor", "Zona", "Quadra"])["Valor"]
                      .agg(lambda v: v.iloc[0] if v.nunique() == 1 else pd.NA))
        for i in df.index[sem_pgv]:
            v = val_quadra.get(tuple(df.loc[i, ["Distrito", "Setor", "Zona", "Quadra"]]), pd.NA)
            msg = (f"secao {df.at[i, 'SecaoPgv']} fora da PGV: usado valor unico da quadra {v}"
                   if pd.notna(v) else f"secao {df.at[i, 'SecaoPgv']} fora da PGV: sem valor de terreno")
            if pd.notna(v):
                df.at[i, "ValorPgv"] = v
            df.at[i, "Observacoes"] = "; ".join(filter(None, [df.at[i, "Observacoes"], msg]))
        print(f"ATENCAO: {sem_pgv.sum()} imovel(is) com secao fora da PGV (ver Observacoes)")

    df["ValorPgv"] = para_numero(df["ValorPgv"])
    df["AreaTerreno"] = para_numero(df["Area_Terreno"])
    df["VVT"] = df["AreaTerreno"] * df["ValorPgv"]

    # VVE-U - PGP
    pgp = pd.read_excel("pgp.xlsx", dtype=str)
    kp = ["KMaterial", "KEdificacao", "KPadrao"]
    for kk, orig_df, orig_pgp in zip(kp, ["Tipo_de_Material", "Tipo_de_Edificacao", "Padrao_Digitado"],
                                     ["TipoDeMaterial", "TipoDeEdificacao", "PadraoDigitado"]):
        df[kk] = df[orig_df].map(chave)
        pgp[kk] = pgp[orig_pgp].map(chave)
    pgp_u = pgp[kp + ["Valor"]].drop_duplicates(subset=kp).rename(columns={"Valor": "VM2E"})
    df = df.merge(pgp_u, on=kp, how="left", validate="m:1")
    df["VM2E"] = para_numero(df["VM2E"])
    df["AreaUnidade"] = para_numero(df["Area_Construida_Unidade_InLoco"])
    if tipo == "territorial":
        df["AreaUnidade"] = np.nan
        edificado = para_numero(df["Valor_Venal_Edificacao_Atual"]).fillna(0) > 0
        for i in df.index[edificado]:
            msg = "planilha territorial, mas a Betha tem valor venal de edificacao no ano atual"
            df.at[i, "Observacoes"] = "; ".join(filter(None, [df.at[i, "Observacoes"], msg]))
    df["VVE-U"] = df["AreaUnidade"] * df["VM2E"]

    sem_pgp = (df["AreaUnidade"].fillna(0) > 0) & df["VM2E"].isna()
    for i in df.index[sem_pgp]:
        msg = (f"unidade {df.at[i, 'Unidade']} sem m2 na PGP "
               f"({df.at[i, 'KMaterial']}/{df.at[i, 'KEdificacao']}/{df.at[i, 'KPadrao']})")
        df.at[i, "Observacoes"] = "; ".join(filter(None, [df.at[i, "Observacoes"], msg]))
    if sem_pgp.any():
        print(f"ATENCAO: {sem_pgp.sum()} unidade(s) com area e sem valor na PGP (ver Observacoes)")

    # Agregados por imovel
    agg = df.groupby("ChaveImovel").agg(VVTE=("VVE-U", "sum"), AreaEdificadaSimulada=("AreaUnidade", "sum"),
                                        Qtd_Unidades=("ID", "size"))
    df = df.drop(columns=["Qtd_Unidades"]).join(agg, on="ChaveImovel")
    pri = df["Principal"]
    df.loc[~pri, ["VVTE", "AreaEdificadaSimulada", "Qtd_Unidades"]] = np.nan
    obs_imovel = df.groupby("ChaveImovel")["Observacoes"].agg(lambda s: "; ".join(x for x in s if x))
    df["ObsImovel"] = np.where(pri, df["ChaveImovel"].map(obs_imovel), "")
    df["VV_T"] = np.where(pri, df["VVT"] + df["VVTE"].fillna(0), np.nan)

    # Aliquota e imposto
    aliq = pd.read_excel("aliquota.xlsx", dtype=str)
    aliq["uso"] = aliq["uso"].map(chave)
    for c in ("min", "max", "aliquota"):
        aliq[c] = pd.to_numeric(aliq[c], errors="coerce")
    df["UsoSimulacao"] = df["Uso_da_Edificacao"].map(chave)
    df.loc[pri & (df["VVTE"].fillna(0) <= 0), "UsoSimulacao"] = "NAO EDIFICADO"

    def achar(uso, vv):
        if pd.isna(vv):
            return np.nan
        f = aliq[aliq["uso"] == uso]
        if f.empty:
            f = aliq[aliq["uso"] == "OUTROS"]
        m = f[(f["min"] <= vv) & (vv <= f["max"])]
        return m["aliquota"].iloc[0] if not m.empty else np.nan

    df["Aliquota"] = [achar(u, v) if p else np.nan
                      for u, v, p in zip(df["UsoSimulacao"], df["VV_T"], pri)]
    df["VI"] = df["VV_T"] * df["Aliquota"] / 100.0
    return df


# =========================================================
# Saida
# =========================================================
# Nome no modelo -> coluna interna (Codigo_Principal agora e coluna propria)
ALIAS_SAIDA = {}


def layout_saida(modelo):
    """Cabecalho da aba SCRIPT_IPTU_SAIDA do modelo; sem modelo, usa COLUNAS_SAIDA."""
    try:
        return list(pd.read_excel(modelo, sheet_name="SCRIPT_IPTU_SAIDA", nrows=0).columns)
    except Exception as e:  # noqa: BLE001 - modelo ausente/sem a aba: segue com o padrao
        print(f"ATENCAO: layout de saida do modelo indisponivel ({e}); usando o padrao")
        return [c for c in COLUNAS_SAIDA if c not in ("Qtd_Unidades", "Observacoes")]


def gravar(df, caminho, layout):
    df["Valor_Venal_Terreno_Simulacao"] = df["VVT"]
    df["Valor_Venal_Edificacao_Simulacao"] = df["VVTE"]
    df["Valor_Venal_Total_Imovel_Simulacao"] = df["VV_T"]
    df["Valor_Aliquota_Simulacao"] = df["Aliquota"]
    df["Valor_Imposto_Simulado"] = df["VI"]
    df["Valor_Impacto"] = df["VI"] - para_numero(df["Valor_Imposto_Atual"])

    valores = [c for c in COLUNAS_SAIDA if c.startswith("Valor_") and "Aliquota" not in c]
    areas = [c for c in COLUNAS_SAIDA if c.startswith("Area_")]
    aliqs = [c for c in COLUNAS_SAIDA if c.startswith("Valor_Aliquota")]

    pri = df.loc[df["Principal"]].sort_values("Inscricao_imobiliaria").reset_index(drop=True)
    for c in valores + areas + aliqs + ["Qtd_Unidades"]:
        pri[c] = para_numero(pri[c])
    for c in valores + areas:
        pri[c] = pri[c].round(2)

    # Aba de saida no layout EXATO do modelo (cabecalho, ordem e grafia)
    saida = pd.DataFrame({col: pri[ALIAS_SAIDA.get(col.strip(), col.strip())]
                          if ALIAS_SAIDA.get(col.strip(), col.strip()) in pri.columns
                          else np.nan for col in layout})
    conferencia = pri[["Codigo", "Codigo_Principal", "ID", "ID_Principal", "Inscricao_imobiliaria",
                       "Qtd_Unidades", "Valor_Imposto_Atual", "Valor_Imposto_Simulado",
                       "Valor_Impacto", "ObsImovel"]].rename(columns={"ObsImovel": "Observacoes"})

    unid = df[["ChaveImovel", "Principal", "Origem", "Codigo", "Codigo_Principal", "ID", "ID_Principal",
               "Inscricao_imobiliaria", "Unidade", "Secao", "SecaoPgv", "ValorPgv", "AreaTerreno",
               "VVT", "KMaterial", "KEdificacao", "KPadrao", "VM2E", "AreaUnidade", "VVE-U",
               "VVTE", "VV_T", "UsoSimulacao", "Aliquota", "VI", "Campos_Betha", "Observacoes"]].copy()
    unid["Principal"] = np.where(unid["Principal"], "SIM", "")
    unid["_o"] = pd.to_numeric(unid["Unidade"], errors="coerce")
    unid = unid.sort_values(["ChaveImovel", "_o"]).drop(columns="_o")

    with pd.ExcelWriter(caminho, engine="openpyxl") as xl:
        saida.to_excel(xl, index=False, sheet_name="SCRIPT_IPTU_SAIDA")
        conferencia.to_excel(xl, index=False, sheet_name="Conferencia")
        unid.to_excel(xl, index=False, sheet_name="Unidades")
        for aba, fmt in (("SCRIPT_IPTU_SAIDA", {**{c: "#,##0.00" for c in valores + areas},
                                                **{c: "0.00" for c in aliqs}}),
                         ("Unidades", {c: "#,##0.00" for c in ["ValorPgv", "AreaTerreno", "VVT", "VM2E",
                                                               "AreaUnidade", "VVE-U", "VVTE", "VV_T", "VI"]})):
            ws = xl.sheets[aba]
            cab = {cel.value: cel.column_letter for cel in ws[1]}
            for nome, mascara in fmt.items():
                if nome in cab:
                    for cel in ws[cab[nome]][1:]:
                        cel.number_format = mascara
                    ws.column_dimensions[cab[nome]].width = max(len(nome) + 2, 14)
    return saida


def main():
    ap = argparse.ArgumentParser(description="Simulacao de IPTU")
    ap.add_argument("--tipo", required=True, choices=TIPOS)
    ap.add_argument("--entrada", default="Calculo IPTU_CAMPOS_SCRIPT.xlsx")
    ap.add_argument("--saida")
    ap.add_argument("--modelo", default="Calculo IPTU_CAMPOS_SCRIPT.xlsx",
                    help="planilha cuja aba SCRIPT_IPTU_SAIDA define o layout da saida")
    ap.add_argument("--ano", type=int, default=2026, help="exercicio 'Atual' (Anterior = ano-1)")
    ap.add_argument("--env", default=".env")
    ap.add_argument("--sem-api", action="store_true", help="nao consulta a Betha (territorial/completa)")
    a = ap.parse_args()
    saida_arq = a.saida or f"planilhaSimulacao_{a.tipo}.xlsx"

    df = ler_entrada(a.entrada)
    print(f"Entrada: {len(df)} linha(s) de {a.entrada} | tipo {a.tipo}")
    if a.sem_api:
        if a.tipo in COM_UNIDADES_BETHA:
            sys.exit(f"--sem-api nao se aplica a {a.tipo}: as unidades vem da Betha")
    else:
        try:
            api = Betha(a.env)
        except RuntimeError as e:
            sys.exit(f"{e} (use --sem-api para territorial/completa)")
        df = completar_com_betha(df, a.tipo, api, a.ano)

    df = calcular(df, a.tipo)
    saida = gravar(df, saida_arq, layout_saida(a.modelo))
    sem_vi = saida["Valor_Imposto_Simulado"].isna().sum()
    print(f"Arquivo gerado: {saida_arq} | {len(saida)} imovel(is), {len(df)} unidade(s)"
          + (f" | ATENCAO: {sem_vi} imovel(is) sem imposto simulado (ver Observacoes)" if sem_vi else ""))


if __name__ == "__main__":
    main()

"""Leitura de imoveis e campos adicionais na API REST da Betha (somente GET).

Credenciais no .env (mesmos nomes dos outros projetos do repositorio):
    tokenAccess=<uuid do token>      (o prefixo "Bearer " e montado aqui)
    userAccess=<user-access>
    (aceita tambem TOKEN_ACCESS / USER_ACCESS)

Rotas usadas (testadas em 19/11/2026):
    GET /dados/v1/imoveis                          filter: id in (...), codigo in (...),
                                                   idImovelPrincipal in (...)
    GET /integracoes-bi/v1/imoveis/campos-adicionais  filter: idImovel in (...) and
                                                   campoAdicional.id in (...) [and ano in (...)]

Na Betha cada unidade e um imovel proprio; as subordinadas apontam para a principal
em idImovelPrincipal (a principal tem idImovelPrincipal = null).
"""
import hashlib
import json
import time
from pathlib import Path

import requests

BASE = "https://tributos.suite.betha.cloud"
LOTE = 100  # ids por requisicao

# Coluna da planilha -> ids do campo adicional, em ordem de preferencia.
# Vale o valor do MAIOR exercicio <= ano da simulacao (o cadastro nao e regravado
# todo ano: em geral o valor vigente esta em 2023 ou antes).
CAMPOS_CADASTRO = {
    "Area_Terreno": (1062953, 1062934),            # Area Terreno; fallback Area Lote
    "Area_Construida_Unidade_InLoco": (1062951,),  # Area Construida Unidade (cadastral)
    "Tipo_de_Material": (1063000,),
    "Tipo_de_Edificacao": (1062962,),
    "Padrao_Digitado": (1062945,),
    "Uso_da_Edificacao": (1063034,),
    "Numero_de_Pavimentos": (1063021,),
    "Taxacao": (1063038,),
    "Tipo_de_Imposto": (1062974,),
    "Classificacao": (1062925,),
    "Patrimonio": (1063030,),
    "Area_Edificada_Total": (1062972,),            # Area Total Construida (por exercicio)
    "Atualizacao_Aerolevantamento": (1659789,),
}

# Valores calculados pela Betha por exercicio -> <base>_Anterior (ano-1) e <base>_Atual (ano)
CAMPOS_ANO = {
    "Valor_Venal_Terreno": 1062910,
    "Valor_Venal_Edificacao": 1062915,
    "Valor_Venal_Total_Imovel": 1062931,
    "Valor_Aliquota": 1062935,
    "Valor_Imposto": 1062940,
}


def ler_env(caminho):
    env = {}
    p = Path(caminho)
    if not p.exists():
        return env
    for linha in p.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if linha and not linha.startswith("#") and "=" in linha:
            k, v = linha.split("=", 1)
            env[k.strip()] = v.strip().strip('"').strip("'")
    return env


def _valor(reg):
    for campo in ("vlCampo", "opcoes", "texto", "areaTexto", "dhCampo"):
        v = reg.get(campo)
        if v not in (None, ""):
            return v
    return None


def _lotes(itens, tamanho=LOTE):
    itens = list(itens)
    for i in range(0, len(itens), tamanho):
        yield itens[i:i + tamanho]


class Betha:
    def __init__(self, env_path, cache_dir="cache_betha"):
        env = ler_env(env_path)
        token = (env.get("tokenAccess") or env.get("TOKEN_ACCESS") or "").replace("Bearer ", "").strip()
        user = (env.get("userAccess") or env.get("USER_ACCESS") or "").strip()
        if not token or not user:
            raise RuntimeError(f"tokenAccess/userAccess ausentes em {env_path}")
        self.sessao = requests.Session()
        self.sessao.headers.update({"Authorization": f"Bearer {token}",
                                    "User-Access": user, "Accept": "application/json"})
        self.cache = Path(cache_dir)
        self.cache.mkdir(exist_ok=True)

    # ---------- HTTP com cache em disco e retentativa ----------
    def _get(self, path, params):
        chave = hashlib.sha1(json.dumps([path, params], sort_keys=True).encode()).hexdigest()
        arq = self.cache / f"{chave}.json"
        if arq.exists():
            return json.loads(arq.read_text(encoding="utf-8"))
        for tentativa in range(3):
            try:
                r = self.sessao.get(BASE + path, params=params, timeout=120)
            except requests.RequestException:
                if tentativa == 2:
                    raise
                time.sleep(2 * (tentativa + 1))
                continue
            if r.status_code >= 500 and tentativa < 2:
                time.sleep(2 * (tentativa + 1))
                continue
            if not r.ok:
                raise RuntimeError(f"GET {path} -> HTTP {r.status_code}: {r.text[:200]}")
            dados = r.json()
            arq.write_text(json.dumps(dados, ensure_ascii=False), encoding="utf-8")
            return dados

    def listar(self, path, filtro, limit=999):
        offset = 0
        while True:
            d = self._get(path, {"filter": filtro, "limit": limit, "offset": offset})
            yield from d.get("content", [])
            if not d.get("hasNext"):
                return
            offset += limit

    # ---------- imoveis ----------
    def imoveis(self, campo, valores, extra=""):
        """campo: 'id', 'codigo' ou 'idImovelPrincipal'."""
        saida = []
        for lote in _lotes(sorted(set(valores))):
            filtro = f"{campo} in ({','.join(map(str, lote))}){extra}"
            saida += list(self.listar("/dados/v1/imoveis", filtro))
        return saida

    def unidades_subordinadas(self, ids_principais):
        return self.imoveis("idImovelPrincipal", ids_principais, " and situacao = 'ATIVADO'")

    # ---------- campos adicionais ----------
    def campos_adicionais(self, ids_imoveis, ano):
        """Retorna {idImovel: {coluna: valor}} com CAMPOS_CADASTRO e CAMPOS_ANO."""
        path = "/integracoes-bi/v1/imoveis/campos-adicionais"
        ids_cad = sorted({i for ids in CAMPOS_CADASTRO.values() for i in ids})
        ids_ano = sorted(CAMPOS_ANO.values())
        brutos = {}  # (idImovel, idCampo) -> {ano: valor}
        for lote in _lotes(sorted(set(ids_imoveis))):
            base = f"idImovel in ({','.join(map(str, lote))})"
            consultas = [
                f"{base} and campoAdicional.id in ({','.join(map(str, ids_cad))}) and ano <= {ano}",
                f"{base} and campoAdicional.id in ({','.join(map(str, ids_ano))}) "
                f"and ano in ({ano - 1},{ano})",
            ]
            for filtro in consultas:
                for reg in self.listar(path, filtro):
                    v = _valor(reg)
                    if v is not None:
                        k = (reg["idImovel"], reg["campoAdicional"]["id"])
                        brutos.setdefault(k, {})[reg.get("ano") or 0] = v

        resultado = {}
        for id_imovel in set(ids_imoveis):
            linha = {}
            for coluna, ids in CAMPOS_CADASTRO.items():
                for id_campo in ids:
                    anos = brutos.get((id_imovel, id_campo))
                    if anos:
                        linha[coluna] = anos[max(anos)]
                        break
            for base, id_campo in CAMPOS_ANO.items():
                anos = brutos.get((id_imovel, id_campo), {})
                if ano - 1 in anos:
                    linha[f"{base}_Anterior"] = anos[ano - 1]
                if ano in anos:
                    linha[f"{base}_Atual"] = anos[ano]
            resultado[id_imovel] = linha
        return resultado


def linha_do_imovel(im):
    """Converte o registro de /dados/v1/imoveis para as colunas da planilha."""
    resp = im.get("responsavel") or {}
    return {
        "ID": str(im["id"]),
        "ID_Principal": str(im.get("idImovelPrincipal") or im["id"]),
        "Codigo": str(im.get("codigo") or ""),
        "Inscricao_imobiliaria": im.get("inscricaoImobiliariaFormatada"),
        "Unidade": str(im.get("unidade") or ""),
        "Secao": str(im.get("secao") or ""),
        "Bairro": (im.get("bairro") or {}).get("nome"),
        "Tipo": (im.get("tipoImovel") or {}).get("descricao"),
        "Situacao": (im.get("situacao") or {}).get("descricao"),
        "Nome_Proprietario": resp.get("nome"),
        "Cpf_Cnpj": resp.get("cpfCnpj"),
    }

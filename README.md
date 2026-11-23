# simulacao-iptu

Simulador de IPTU em Python. A partir de uma planilha de atualização cadastral
(aerolevantamento, novo padrão construtivo, conversão para territorial), calcula o
**valor venal e o imposto simulados** de cada imóvel e compara com o imposto
lançado no sistema tributário (**Betha Tributos**).

> Este repositório contém **apenas o código**. As planilhas de entrada e as tabelas de
> valores (PGV, PGP e alíquotas) não são versionadas: as entradas têm dados de
> contribuintes (nome/CPF). O formato de cada planilha está descrito abaixo.

---

## Sumário

1. [Como funciona](#1-como-funciona)
2. [Estrutura da pasta](#2-estrutura-da-pasta)
3. [Instalação](#3-instalação)
4. [Planilhas obrigatórias (tabelas de valores)](#4-planilhas-obrigatórias-tabelas-de-valores)
5. [Os tipos de planilha de entrada](#5-os-tipos-de-planilha-de-entrada)
6. [Formato da planilha de entrada](#6-formato-da-planilha-de-entrada)
7. [Rodando](#7-rodando)
8. [Resultado](#8-resultado)
9. [Fórmulas](#9-fórmulas)
10. [Mensagens de atenção](#10-mensagens-de-atenção)
11. [Cuidados](#11-cuidados)

---

## 1. Como funciona

```
planilha de entrada ──► completa vazios pela API da Betha ──► calcula VVT, VVE, alíquota, imposto ──► xlsx de resultado
  (aerolevantamento)       (somente leitura, opcional)           (PGV, PGP, alíquotas)                 (3 abas)
```

- **A planilha prevalece sobre a Betha.** A Betha só completa o que estiver vazio e
  traz os valores lançados (Anterior/Atual) para comparação.
- A API é usada **somente para leitura** (GET). Nada é gravado no sistema.
- Na Betha, cada unidade é um imóvel próprio; as unidades subordinadas apontam para a
  principal em `idImovelPrincipal`.

---

## 2. Estrutura da pasta

```
simulacao-iptu/
├── simulacaoIPTU.py                 # simulador (entrada, cálculo, saída)
├── betha_api.py                     # leitura na API REST da Betha (somente GET)
├── requirements.txt
├── .env.example                     # modelo de credenciais -> copiar para .env
├── LEIA-ME.txt                      # metodologia de correção (IPCA) e arredondamento
├── pgv.xlsx                         # (não versionado) Planta Genérica de Valores - terreno
├── pgp.xlsx                         # (não versionado) Planta Genérica de valores Prediais - edificação
├── aliquota.xlsx                    # (não versionado) faixas de alíquota por uso
├── Calculo IPTU_CAMPOS_SCRIPT.xlsx   # (não versionado, opcional) modelo do layout de saída
└── entradas/                        # (não versionado) planilhas a simular
    ├── 1_territorial.xlsx
    ├── 2_novo_padrao.xlsx
    └── 3_edificacoes.xlsx
```

`pgv.xlsx`, `pgp.xlsx` e `aliquota.xlsx` precisam estar na **pasta de onde o
comando é executado**.

---

## 3. Instalação

Requer **Python 3.10+**.

```bash
pip install -r requirements.txt
```

Credenciais da API (opcional, mas necessário para resultado completo):

```bash
cp .env.example .env
```

```
tokenAccess=<uuid do token, sem "Bearer ">
userAccess=<user-access>
```

O prefixo `Bearer ` é montado pelo código. Também são aceitos os nomes
`TOKEN_ACCESS` / `USER_ACCESS`. **Nunca versione nem envie o `.env`.**

Sem `.env`, use `--sem-api` (somente `territorial` e `completa`). Nesse modo os valores
lançados (Anterior/Atual) não são buscados e as unidades novas in loco não são
agrupadas no imóvel principal, então **o resultado não serve para comparação**.

---

## 4. Planilhas obrigatórias (tabelas de valores)

As três tabelas são lidas da **primeira aba**. Os nomes das colunas devem ser
exatamente estes.

### 4.1 `pgv.xlsx` — valor do m² do terreno

| Coluna | Exemplo | Descrição |
|---|---|---|
| `Distrito` | `00` | 1º segmento da inscrição |
| `Setor` | `01` | 2º segmento |
| `Zona` | `00` | 3º segmento |
| `Quadra` | `00` | 4º segmento |
| `Secao` | `3003` | seção de logradouro (completada com zeros até 5 dígitos) |
| `Nome da Rua` | `PRESIDENTE VARGAS` | informativo |
| `Valor` | `47` | valor do m² do terreno |

A chave é `Distrito + Setor + Zona + Quadra + Secao`. Se a seção do imóvel não estiver
na PGV e a quadra tiver **um único valor**, esse valor é usado (com aviso).

### 4.2 `pgp.xlsx` — valor do m² da edificação

| Coluna | Exemplo |
|---|---|
| `PCODE` | `1.1.3` (informativo) |
| `TipoDeMaterial` | `ALVENARIA` |
| `TipoDeEdificacao` | `CASA RESIDENCIAL` |
| `PadraoDigitado` | `MEDIO` |
| `Valor` | `1299` |

A comparação ignora acentos, hífens, maiúsculas/minúsculas e prefixos numéricos
(`1 - ALVENARIA`, `CASA-RESIDENCIAL` e `MÉDIO` casam com a tabela).

### 4.3 `aliquota.xlsx` — faixas de alíquota

| Coluna | Exemplo | Descrição |
|---|---|---|
| `uso` | `RESIDENCIAL` | uso da edificação; `NAO EDIFICADO` para terreno; `OUTROS` é o padrão quando o uso não está na tabela |
| `min` | `0` | valor venal total mínimo da faixa |
| `max` | `74801` | valor venal total máximo da faixa |
| `aliquota` | `0.5` | em % |

### 4.4 `Calculo IPTU_CAMPOS_SCRIPT.xlsx` — modelo de saída (opcional)

Só o **cabeçalho** da aba `SCRIPT_IPTU_SAIDA` é usado: define a ordem e a grafia das
colunas do resultado. Sem esse arquivo, o simulador usa o layout padrão e mostra um aviso.

---

## 5. Os tipos de planilha de entrada

| Tipo (`--tipo`) | Quando usar | O que vem da planilha | O que vem da Betha |
|---|---|---|---|
| `territorial` | Imóveis que deixaram de ter edificação (predial → territorial) | Cadastro e área do terreno | Campos vazios de cadastro/terreno e valores Anterior/Atual. **Edificação é ignorada**; alíquota da faixa `NAO EDIFICADO` |
| `completa` | A planilha já traz **todas as unidades** do imóvel (aerolevantamento, novo padrão construtivo) | Tudo da construção (material, tipo, padrão, área in loco) | Só campos vazios de cadastro/terreno, `Uso_da_Edificacao` e valores Anterior/Atual |
| `sem_unidades` | A planilha traz **só a unidade principal** | A principal | As **unidades subordinadas** ativas (`idImovelPrincipal`), com material, tipo, padrão, uso e área |
| `novo_padrao` | A planilha traz **só a unidade principal com o padrão construtivo novo** | A principal e o `Padrao_Digitado` novo | Igual a `sem_unidades`; depois o **padrão da principal é aplicado em todas as unidades** do imóvel (o padrão anterior fica em `Observacoes`) |

Exemplo de planilhas reais e o tipo usado:

| Planilha | Tipo | Observação |
|---|---|---|
| `1_territorial` | `territorial` | A inscrição atual é a `Inscricao_imobiliaria_InLoco` (termina em `.0`; terminava em `.1` quando era predial) |
| `2_novo_padrao` | **`novo_padrao`** | Cada linha é a unidade **principal**, com o padrão novo e **sem área construída**. O simulador busca na Betha a área da principal e todas as demais unidades do imóvel, e aplica o padrão novo da principal em **todas as unidades**. Com `completa` o imóvel seria calculado sem edificação (alíquota de terreno) |
| `3_edificacoes` | `completa` | `Inscricao_imobiliaria` é a inscrição englobada da principal e se repete nas unidades; `Inscricao_imobiliaria_InLoco` identifica as **unidades novas** achadas em campo, ainda sem código/ID no sistema |

---

## 6. Formato da planilha de entrada

Aceita `.xlsx` ou `.csv`.

- **xlsx:** lê a aba `SCRIPT_IPTU_ENTRADA`; se ela não existir, lê a **primeira aba**
  (ex.: aba `CGR`).
- **csv:** separador `;`, encoding ISO-8859-1.
- Espaços no início/fim dos nomes das colunas são ignorados. Coluna ausente é tratada
  como vazia. Números aceitam `1234.56`, `1234,56` e `1.234,56`.

### 6.1 Colunas

| Coluna | Uso |
|---|---|
| `ID` | ID do imóvel na Betha. Se vazio, é resolvido pelo `Codigo` |
| `ID_Principal` | ID da unidade principal (vazio ou `0` = é a principal) |
| `Codigo` | Código do imóvel |
| `Codigo_Principal` | Código da principal (recalculado na saída) |
| `Inscricao_imobiliaria` | Inscrição `DD.SS.ZZ.QQ.LLL.UUU.D`; os 4 primeiros segmentos são a chave da PGV |
| `Inscricao_imobiliaria_InLoco` | Inscrição levantada em campo (usada quando a cadastral está vazia) |
| `Unidade` | Número da unidade (ordena e escolhe a principal) |
| `Secao` | Seção de logradouro (chave da PGV) |
| `Tipo_de_Edificacao`, `Tipo_de_Material`, `Padrao_Digitado` | Chave da PGP |
| `Uso_da_Edificacao` | Faixa de alíquota |
| `Area_Construida_Unidade_InLoco` | Área construída da unidade (m²) |
| `Area_Terreno` | Área do terreno (m²), da unidade principal |
| `Valor_Venal_Terreno_Anterior/Atual`, `Valor_Venal_Edificacao_Anterior/Atual`, `Valor_Venal_Total_Imovel_Anterior/Atual`, `Valor_Aliquota_Anterior/Atual`, `Valor_Imposto_Anterior/Atual` | Valores lançados (completados pela Betha se vazios) |
| `Nome_Proprietario`, `Cpf_Cnpj`, `Bairro`, `Tipo`, `Situacao`, `Taxacao`, `Numero_de_Pavimentos`, `Classificacao`, `Patrimonio`, `Tipo_de_Imposto`, `Atualizacao_Aerolevantamento`, `Area_Edificada_Total`, `Area_Construida_Obras`, `Ano` | Informativos, repassados para a saída |
| `Area_Construida_Unidade` | (CSV do relatório cadastral) usada quando `..._InLoco` está vazia |

### 6.2 Como as linhas viram imóveis

1. Linha com `ID`: pertence ao imóvel do seu `ID_Principal` (ou ao próprio ID, se for a principal).
2. Linha **sem `ID` e com `Codigo`**: o ID é resolvido na Betha pelo código.
3. Linha **sem `ID` e sem `Codigo`** (unidade nova levantada em campo): entra no imóvel
   cadastrado que tem a **mesma `Inscricao_imobiliaria`** (inscrição englobada).
4. A unidade principal é sempre uma linha cadastrada (com `ID`), a de menor `Unidade`.

### 6.3 Colunas de principal na saída

| Linha | `Codigo_Principal` | `ID_Principal` |
|---|---|---|
| Imóvel principal | o próprio código | o próprio ID |
| Unidade englobada | código da principal | ID da principal |

---

## 7. Rodando

Execute **dentro da pasta do projeto** (onde estão `pgv/pgp/aliquota`):

```bash
python simulacaoIPTU.py --tipo territorial --entrada entradas/1_territorial.xlsx --saida 1_territorial_SIMULADO.xlsx
python simulacaoIPTU.py --tipo novo_padrao  --entrada entradas/2_novo_padrao.xlsx --saida 2_novo_padrao_SIMULADO.xlsx
python simulacaoIPTU.py --tipo completa    --entrada entradas/3_edificacoes.xlsx --saida 3_edificacoes_SIMULADO.xlsx
```

| Parâmetro | Padrão | Descrição |
|---|---|---|
| `--tipo` | (obrigatório) | `territorial`, `completa`, `sem_unidades` ou `novo_padrao` |
| `--entrada` | `Calculo IPTU_CAMPOS_SCRIPT.xlsx` | Planilha ou CSV de entrada |
| `--saida` | `planilhaSimulacao_<tipo>.xlsx` | Arquivo gerado |
| `--ano` | `2026` | Exercício "Atual" (Anterior = ano − 1) |
| `--modelo` | `Calculo IPTU_CAMPOS_SCRIPT.xlsx` | Modelo do layout de saída |
| `--env` | `.env` | Arquivo de credenciais |
| `--sem-api` | — | Não consulta a Betha (não vale para `sem_unidades` e `novo_padrao`) |

Saída esperada:

```
Entrada: 2359 linha(s) de entradas/3_edificacoes.xlsx | tipo completa
Arquivo gerado: 3_edificacoes_SIMULADO.xlsx | 824 imovel(is), 2359 unidade(s)
```

---

## 8. Resultado

O `.xlsx` gerado tem três abas:

| Aba | Conteúdo |
|---|---|
| `SCRIPT_IPTU_SAIDA` | **Uma linha por imóvel principal** (os valores somam todas as unidades), no layout do modelo: dados cadastrais, valores Anterior/Atual/Simulação, `Valor_Imposto_Simulado` e `Valor_Impacto` (simulado − atual) |
| `Conferencia` | Resumo por imóvel: `Codigo`, `Codigo_Principal`, `ID`, `ID_Principal`, `Qtd_Unidades`, imposto atual × simulado, impacto e `Observacoes` |
| `Unidades` | Cálculo passo a passo, uma linha por unidade: `ValorPgv`, `VVT`, chaves da PGP (`KMaterial`, `KEdificacao`, `KPadrao`), `VM2E`, `AreaUnidade`, `VVE-U`, `VVTE`, `VV_T`, `UsoSimulacao`, `Aliquota`, `VI`, `Origem` (PLANILHA/BETHA) e `Campos_Betha` (o que foi completado pela API) |

---

## 9. Fórmulas

```
VVT    = Area_Terreno (principal) × m² da PGV (Distrito/Setor/Zona/Quadra/Seção)
VVE-U  = Area_Construida_Unidade_InLoco × m² da PGP (material/tipo/padrão)
VVTE   = Σ VVE-U das unidades do imóvel
VV_T   = VVT + VVTE
Aliq   = faixa de aliquota.xlsx por Uso_da_Edificacao da principal e VV_T
         (VVTE = 0 ou tipo territorial -> NAO EDIFICADO)
VI     = VV_T × Aliq / 100
Impacto = VI − Valor_Imposto_Atual
```

Metodologia de correção dos valores (IPCA) e regra de arredondamento: ver `LEIA-ME.txt`.

---

## 10. Mensagens de atenção

| Mensagem | Significado / ação |
|---|---|
| `N ID(s) nao encontrados na Betha` | IDs da planilha que não existem no sistema — conferir a planilha |
| `secao X fora da PGV` | Seção sem valor na PGV; usado o valor único da quadra, se houver. Ver `Observacoes` |
| `unidade N sem m2 na PGP (material/tipo/padrao)` | Combinação inexistente na PGP — a unidade fica sem valor de edificação (ex.: piscina sem material cadastrado na Betha) |
| `planilha territorial, mas a Betha tem valor venal de edificacao` | Imóvel tratado como territorial ainda tem edificação lançada no sistema |
| `Secao/Inscricao planilha X x Betha Y` | Divergência entre planilha e cadastro (a planilha prevalece) |
| `layout de saida do modelo indisponivel` | Sem `Calculo IPTU_CAMPOS_SCRIPT.xlsx`; usado o layout padrão |

---

## 11. Cuidados

- **Cache da API:** respostas ficam em `cache_betha/` e **não expiram**. Se o cadastro
  mudou no sistema, apague a pasta antes de rodar.
- **Dados pessoais:** entradas e resultados contêm nome e CPF. O `.gitignore` bloqueia
  `*.xlsx`, `*.csv` e `entradas/`; não force o commit.
- **Novo exercício:** atualize `pgv.xlsx`, `pgp.xlsx` e `aliquota.xlsx` e use `--ano`.
- Os IDs de campos adicionais consultados na Betha estão em `betha_api.py`
  (`CAMPOS_CADASTRO`, `CAMPOS_ANO`) e são específicos do município; ajuste-os para
  outro ambiente.

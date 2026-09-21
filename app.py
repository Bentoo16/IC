import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import re
import unicodedata

# ---------------------------------------------------------------------------
# CSS customizado
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .tabela-respostas {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        font-size: 0.85rem;
        margin-bottom: 1.5rem;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }
    .tabela-respostas th {
        background: #1F4E79;
        color: #ffffff;
        padding: 0.6rem 0.5rem;
        text-align: center;
        border: 1px solid #163a5c;
        font-weight: 600;
    }
    .tabela-respostas th.caso-header {
        background: #2E75B6;
        color: #ffffff;
    }
    .tabela-respostas tr:nth-child(even) td {
        background: #F4F7FB;
    }
    .tabela-respostas tr:hover td {
        background: #E8F0FA;
    }
    .tabela-respostas td {
        padding: 0.45rem 0.5rem;
        border: 1px solid #dee2e6;
        text-align: center;
        transition: background 0.15s ease-in-out;
    }
    .tabela-respostas td:first-child {
        text-align: left;
        font-weight: 500;
        color: #1F1F1F;
    }
    .preview-box {
        background: #F4F7FB;
        border-left: 5px solid #1F4E79;
        padding: 1rem 1.2rem;
        border-radius: 8px;
        margin-bottom: 1.2rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        line-height: 1.6;
    }
    .preview-box strong {
        color: #1F4E79;
    }
    .progresso-texto {
        font-size: 0.85rem;
        color: #555;
        margin-bottom: 0.3rem;
    }
    .stButton>button {
        font-weight: 600;
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Funções auxiliares
# ---------------------------------------------------------------------------
def extrair_numero(nome):
    match = re.search(r'\d+', nome)
    return int(match.group()) if match else 0


def gerar_tabela_html(casos_ordenados, perguntas_ordenadas, escolhas_casos):
    html = "<table class='tabela-respostas'>"
    html += "<tr><th rowspan='2'>Pergunta</th>"
    for caso in casos_ordenados:
        html += f"<th class='caso-header' colspan='2'>{caso}</th>"
    html += "</tr><tr>"
    for _ in casos_ordenados:
        html += "<th>Sim</th><th>Não</th>"
    html += "</tr>"
    for pergunta in perguntas_ordenadas:
        html += "<tr>"
        html += f"<td>{pergunta}</td>"
        for caso in casos_ordenados:
            item = escolhas_casos.get(caso, {}).get(pergunta, {})
            resposta = item.get("resposta", "-") if isinstance(item, dict) else item
            sim_cell = "X" if resposta == "Sim" else ""
            nao_cell = "X" if resposta == "Não" else ""
            html += f"<td>{sim_cell}</td><td>{nao_cell}</td>"
        html += "</tr>"
    html += "</table>"
    return html


def contar_perguntas(grupos):
    total = 0
    for qs in grupos.values():
        total += len(qs)
    return total


def sanitizar_nome_arquivo(texto):
    """Transforma um texto livre (ex.: nome do serviço) em algo seguro para
    usar como nome de arquivo: sem acentos, sem caracteres especiais e com
    espaços trocados por underscore."""
    if not texto or not texto.strip():
        return "Servico"
    texto_sem_acento = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    texto_limpo = re.sub(r'[^\w\s-]', '', texto_sem_acento).strip()
    texto_final = re.sub(r'[-\s]+', '_', texto_limpo)
    return texto_final or "Servico"


def set_cell_shading(cell, color):
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    shading.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shading)


# ---------------------------------------------------------------------------
# Paleta e helpers visuais do documento .docx
# ---------------------------------------------------------------------------
COR_PRIMARIA = "1F4E79"   # azul escuro - cabeçalho principal das tabelas
COR_SECUNDARIA = "2E75B6"  # azul médio - linha "Caso" das tabelas
COR_ZEBRA = "F4F7FB"       # cinza-azulado bem claro - linhas alternadas


def estilizar_documento(doc):
    """Define fonte base e cores dos títulos para todo o documento."""
    estilo_normal = doc.styles["Normal"]
    estilo_normal.font.name = "Calibri"
    estilo_normal.font.size = Pt(11)

    tamanhos_titulo = {0: 20, 1: 15, 2: 13}
    for nivel, tamanho in tamanhos_titulo.items():
        estilo = doc.styles[f"Heading {nivel}" if nivel > 0 else "Title"]
        estilo.font.name = "Calibri"
        estilo.font.size = Pt(tamanho)
        estilo.font.color.rgb = RGBColor.from_string(COR_PRIMARIA)
        estilo.font.bold = True


def formatar_celula(cell, texto=None, cor_fundo=None, cor_texto=None, negrito=False, centralizar=True):
    """Preenche uma célula de tabela com texto formatado (cor, negrito, fundo)."""
    if texto is not None:
        cell.text = ""
        paragrafo = cell.paragraphs[0]
        run = paragrafo.add_run(texto)
    else:
        paragrafo = cell.paragraphs[0]
        run = paragrafo.runs[0] if paragrafo.runs else paragrafo.add_run("")
    run.bold = negrito
    if cor_texto:
        run.font.color.rgb = RGBColor.from_string(cor_texto)
    if centralizar:
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if cor_fundo:
        set_cell_shading(cell, cor_fundo)


# ---------------------------------------------------------------------------
# Configuração da IA
# ---------------------------------------------------------------------------
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-3.1-flash-lite')

st.title("Gerador do Instrumento para Análise de Qualidade em Mamografia")

# ---------------------------------------------------------------------------
# Inicialização do session_state
# ---------------------------------------------------------------------------
if "dados_cabecalho" not in st.session_state:
    st.session_state.dados_cabecalho = {
        "mamografo_fabricante": "",
        "mamografo_modelo": "",
        "cnes": "",
        "qiid": "",
        "tipo_mamografo": None,
        "instituicao": "",
        "cidade": "",
        "estado": "",
        "servico": "",
    }
if "servico" not in st.session_state.dados_cabecalho:
    st.session_state.dados_cabecalho["servico"] = ""
if "casos_salvos" not in st.session_state:
    st.session_state.casos_salvos = {}
if "relatorios_ia" not in st.session_state:
    st.session_state.relatorios_ia = {}
if "relatorio_geral_salvo" not in st.session_state:
    st.session_state.relatorio_geral_salvo = None
if "consideracoes_caso" not in st.session_state:
    st.session_state.consideracoes_caso = {}
if "consideracoes_gerais" not in st.session_state:
    st.session_state.consideracoes_gerais = ""
if "escolhas_casos" not in st.session_state:
    st.session_state.escolhas_casos = {}
if "identificacao_exames" not in st.session_state:
    st.session_state.identificacao_exames = {}
if "dados_adicionais_casos" not in st.session_state:
    st.session_state.dados_adicionais_casos = {}

# ---------------------------------------------------------------------------
# Barra de progresso da sessão
# ---------------------------------------------------------------------------
total_casos = 5
casos_feitos = len(st.session_state.casos_salvos)
if casos_feitos > 0:
    st.markdown(f"<div class='progresso-texto'>Progresso: {casos_feitos} de {total_casos} casos analisados</div>", unsafe_allow_html=True)
    st.progress(casos_feitos / total_casos)

# ---------------------------------------------------------------------------
# Cabeçalho - dados da instituição
# ---------------------------------------------------------------------------
st.markdown("---")
st.subheader("Dados do Cabeçalho")

col1, col2 = st.columns(2)
with col1:
    fabricante = st.text_input("Mamógrafo - Fabricante:", value=st.session_state.dados_cabecalho["mamografo_fabricante"])
with col2:
    modelo = st.text_input("Mamógrafo - Modelo:", value=st.session_state.dados_cabecalho["mamografo_modelo"])

cnes = st.text_input("CNES:", value=st.session_state.dados_cabecalho["cnes"])
qiid = st.text_input("QIID:", value=st.session_state.dados_cabecalho["qiid"])

tipo = st.radio(
    "Tipo de mamógrafo:",
    ["Convencional", "Digital CR", "Digital DR", "DR retrofit"],
    index=0
    if st.session_state.dados_cabecalho["tipo_mamografo"] is None
    else ["Convencional", "Digital CR", "Digital DR", "DR retrofit"].index(
        st.session_state.dados_cabecalho["tipo_mamografo"]
    ),
    horizontal=True,
    key="tipo_mamografo_radio",
)

instituicao = st.text_input("Instituição:", value=st.session_state.dados_cabecalho["instituicao"])

col3, col4 = st.columns(2)
with col3:
    cidade = st.text_input("Cidade:", value=st.session_state.dados_cabecalho["cidade"])
with col4:
    estado = st.text_input("Estado:", value=st.session_state.dados_cabecalho["estado"])

servico = st.text_input("Serviço:", value=st.session_state.dados_cabecalho["servico"])

st.session_state.dados_cabecalho = {
    "mamografo_fabricante": fabricante,
    "mamografo_modelo": modelo,
    "cnes": cnes,
    "qiid": qiid,
    "tipo_mamografo": tipo,
    "instituicao": instituicao,
    "cidade": cidade,
    "estado": estado,
    "servico": servico,
}

# ---------------------------------------------------------------------------
# Biblioteca de perguntas (organizada por grupos)
# ---------------------------------------------------------------------------
perguntas = {
    "Avaliação dos Critérios de Laudos": {
        "Resumo da história presente": {
            "opcoes": {"Sim": " ", "Não": "É importante que nos laudos conste a indicação do exame. Essa indicação deve conter uma história resumida da paciente (exame de rastreamento x diagnóstico / história familiar / antecedentes cirúrgicos e resultados de biópsias / sintomas e queixas da paciente ... )."},
        },
        "Utiliza corretamente o Léxico BI-RADS® ou SISMAMA": {
            "opcoes": {"Sim": " ", "Não": "No laudo deste exame não foi utilizado corretamente o léxico do BI-RADS® ou do SISMAMA."},
        },
        "Classifica corretamente o exame segundo o BI-RADS®": {
            "opcoes": {"Sim": " ", "Não": "O exame não foi classificado corretamente."},
        },
        "Recomendação correta segundo o BI-RADS®": {
            "opcoes": {"Sim": " ", "Não": " No laudo enviado para avaliação, o exame não foi classificado corretamente."},
        },
        "Interpretou corretamente todos os achados do exame": {
            "opcoes": {"Sim": "", "Não": " Os achados do exame não foram interpretados corretamente."},
        },
    },
}

# ---------------------------------------------------------------------------
# Perguntas adicionais (texto livre) - não entram no texto enviado à IA,
# servem apenas para compor uma tabela própria no documento .docx.
# ---------------------------------------------------------------------------
perguntas_adicionais_texto = [
    "Padrão de Mama segundo o serviço: (MARCAR: A, B, C ou D)",
    "Tipo de achado segundo o serviço: (MARCAR: MCF, CALC, NOD, DISTORC, ASSIM, etc)",
    "Classificação BI-RADS® do serviço: (MARCAR: 0, 1, 2, 3, 4, 5, 6)",
    "Classificação BI-RADS® dos avaliadores: (MARCAR: 0, 1, 2, 3, 4, 5, 6)",
]

# ---------------------------------------------------------------------------
# Atalho: marcar os 5 casos como "Sim" em tudo (sem considerações específicas)
# ---------------------------------------------------------------------------
st.markdown("---")
st.subheader("Atalho Rápido")
possui_consideracoes = st.radio(
    "Algum dos 5 casos possui consideração específica a ser registrada?",
    ["Sim, vou analisar cada caso individualmente", "Não, os 5 casos estão 'Sim' em tudo"],
    key="modo_atalho_sem_consideracoes",
)

if possui_consideracoes == "Não, os 5 casos estão 'Sim' em tudo":
    nomes_casos_atalho = [f"Caso {n}" for n in range(1, total_casos + 1)]
    casos_ja_salvos_atalho = [nome for nome in nomes_casos_atalho if nome in st.session_state.casos_salvos]
    confirmar_atalho = True
    if casos_ja_salvos_atalho:
        st.warning(f"Isto vai sobrescrever os casos já salvos: {', '.join(casos_ja_salvos_atalho)}.")
        confirmar_atalho = st.checkbox(
            "Confirmo que desejo sobrescrever os casos acima.",
            key="confirmar_atalho_sem_consideracoes",
        )

    if st.button(
        "Marcar os 5 casos como 'Sim' em tudo e salvar",
        type="primary",
        use_container_width=True,
        disabled=not confirmar_atalho,
    ):
        TEXTO_LAUDO_SEM_CONSIDERACOES = "Sem considerações específicas sobre o laudo deste caso."
        titulos_laudo_atalho = list(perguntas["Avaliação dos Critérios de Laudos"].keys())
        for nome in nomes_casos_atalho:
            escolhas_atalho = {
                titulo: {"resposta": "Sim", "sub_opcao": []}
                for titulo in titulos_laudo_atalho
            }
            st.session_state.casos_salvos[nome] = TEXTO_LAUDO_SEM_CONSIDERACOES
            st.session_state.consideracoes_caso[nome] = ""
            st.session_state.identificacao_exames.setdefault(nome, "")
            st.session_state.escolhas_casos[nome] = escolhas_atalho
            st.session_state.relatorios_ia[nome] = TEXTO_LAUDO_SEM_CONSIDERACOES
        st.session_state.docx_bytes = None
        st.success("Os 5 casos foram marcados como 'Sim' em tudo e salvos!")
        st.rerun()

# ---------------------------------------------------------------------------
# Combinações de respostas: quando um CONJUNTO de perguntas específicas
# recebe a mesma resposta (ex.: "Não" em duas perguntas ao mesmo tempo),
# usa-se UMA frase conjunta no lugar das frases individuais de cada
# pergunta. A frase aparece só uma vez, mesmo com várias perguntas
# disparando a mesma combinação.
#
# Para CRIAR uma nova combinação, copie um bloco abaixo e ajuste
# "perguntas" (a lista de títulos que precisam bater), "resposta_gatilho"
# e "texto". Para REMOVER, apague o bloco correspondente da lista.
# ---------------------------------------------------------------------------
COMBINACOES_RESPOSTA = [
    {
        "perguntas": [
            "Classifica corretamente o exame segundo o BI-RADS®",
            "Recomendação correta segundo o BI-RADS®",
        ],
        "resposta_gatilho": "Não",
        "texto": "[TODO: Breno, me diga a frase exata que deve aparecer quando as duas perguntas acima forem 'Não' juntas.]",
    },
]

# ---------------------------------------------------------------------------
# Seleção do caso e perguntas
# ---------------------------------------------------------------------------
st.markdown("---")
caso_atual = st.selectbox("Escolha o Caso que vai analisar agora:", [1, 2, 3, 4, 5])
nome_caso = f"Caso {caso_atual}"

st.markdown("---")
st.subheader("Dados Adicionais do Caso")
valores_adicionais_salvos = st.session_state.dados_adicionais_casos.get(nome_caso, {})
dados_adicionais_temp = {}
for pergunta_extra in perguntas_adicionais_texto:
    dados_adicionais_temp[pergunta_extra] = st.text_input(
        pergunta_extra,
        value=valores_adicionais_salvos.get(pergunta_extra, ""),
        key=f"extra_{pergunta_extra}_c{caso_atual}",
    )

respostas_temporarias = []
abas_grupos = st.tabs(list(perguntas.keys()))
for idx, (nome_grupo, questoes) in enumerate(perguntas.items()):
    with abas_grupos[idx]:
        for titulo, info in questoes.items():
            st.subheader(titulo)
            escolha = st.radio("Selecione:", list(info["opcoes"].keys()), key=f"radio_{titulo}_c{caso_atual}", horizontal=True)
            gatilho = info.get("gatilho_sub_opcoes", "Não")
            sub_escolha = []
            if "sub_opcoes" in info and escolha == gatilho:
                sub_escolha = st.multiselect("Especifique:", list(info["sub_opcoes"].keys()), key=f"sub_{titulo}_c{caso_atual}")
            obs = st.text_input("Considerações adicionais:", key=f"obs_{titulo}_c{caso_atual}", placeholder="Opcional")
            respostas_temporarias.append({"titulo": titulo, "escolha": escolha, "sub_escolha": sub_escolha, "obs": obs})

st.markdown("---")
id_exame = st.text_input(
    "Identificação do Exame:",
    value=st.session_state.identificacao_exames.get(nome_caso, ""),
    key=f"id_exame_c{caso_atual}",
)
consideracoes_caso = st.text_area(
    "Considerações adicionais para este caso (opcional):",
    value=st.session_state.consideracoes_caso.get(nome_caso, ""),
    key=f"consideracoes_c{caso_atual}",
    height=100,
)

caso_ja_existe = nome_caso in st.session_state.casos_salvos
confirmacao = True
if caso_ja_existe:
    st.warning(f"O {nome_caso} já foi salvo anteriormente.")
    confirmacao = st.checkbox("Deseja sobrescrever o relatório existente?", key=f"conf_{caso_atual}")

if st.button(f"Analisar e Salvar {nome_caso}", type="primary", use_container_width=True):
    if caso_ja_existe and not confirmacao:
        st.warning("Marque a confirmação para sobrescrever o caso.")
    else:
        TEXTO_LAUDO_SEM_CONSIDERACOES = "Sem considerações específicas sobre o laudo deste caso."
        titulos_grupo_laudo = set(perguntas["Avaliação dos Critérios de Laudos"].keys())
        respostas_grupo_laudo = {
            item["titulo"]: item["escolha"]
            for item in respostas_temporarias
            if item["titulo"] in titulos_grupo_laudo
        }
        laudo_tudo_sim = (
            len(respostas_grupo_laudo) == len(titulos_grupo_laudo)
            and all(escolha == "Sim" for escolha in respostas_grupo_laudo.values())
        )

        respostas_por_titulo = {item["titulo"]: item["escolha"] for item in respostas_temporarias}
        combinacoes_disparadas = [
            combinacao
            for combinacao in COMBINACOES_RESPOSTA
            if all(
                respostas_por_titulo.get(titulo) == combinacao["resposta_gatilho"]
                for titulo in combinacao["perguntas"]
            )
        ]
        titulos_cobertos_por_combinacao = {
            titulo
            for combinacao in combinacoes_disparadas
            for titulo in combinacao["perguntas"]
        }

        respostas_finais = []
        laudo_texto_inserido = False
        combinacoes_inseridas = set()
        for item in respostas_temporarias:
            if item["titulo"] in titulos_grupo_laudo and laudo_tudo_sim:
                # Todas as perguntas deste grupo foram "Sim": em vez dos textos
                # em branco de cada pergunta, insere uma única frase fixa.
                if not laudo_texto_inserido:
                    respostas_finais.append(TEXTO_LAUDO_SEM_CONSIDERACOES)
                    laudo_texto_inserido = True
                if item["obs"]:
                    respostas_finais.append(f"Detalhe adicional: {item['obs']}")
                continue
            if item["titulo"] in titulos_cobertos_por_combinacao:
                # Esta pergunta faz parte de uma combinação disparada: insere a
                # frase conjunta uma única vez, no lugar da frase individual.
                combinacao = next(
                    c for c in combinacoes_disparadas if item["titulo"] in c["perguntas"]
                )
                chave_combinacao = tuple(combinacao["perguntas"])
                if chave_combinacao not in combinacoes_inseridas:
                    respostas_finais.append(combinacao["texto"])
                    combinacoes_inseridas.add(chave_combinacao)
                if item["obs"]:
                    respostas_finais.append(f"Detalhe adicional: {item['obs']}")
                continue
            # Encontrar o grupo e a pergunta
            for questoes in perguntas.values():
                if item["titulo"] in questoes:
                    info_pergunta = questoes[item["titulo"]]
                    break
            gatilho_pergunta = info_pergunta.get("gatilho_sub_opcoes", "Não")
            if item["escolha"] == gatilho_pergunta and item["sub_escolha"]:
                frase_base = " ".join(
                    info_pergunta["sub_opcoes"][opcao] for opcao in item["sub_escolha"]
                )
            else:
                frase_base = info_pergunta["opcoes"][item["escolha"]]
            if item["obs"]:
                frase_base += f" Detalhe adicional: {item['obs']}"
            respostas_finais.append(frase_base)

        texto_bruto = " ".join(respostas_finais)
        texto_para_ia = texto_bruto
        if consideracoes_caso.strip():
            texto_para_ia += f"\n\n{consideracoes_caso}"

        escolhas = {
            item["titulo"]: {
                "resposta": item["escolha"],
                "sub_opcao": item["sub_escolha"],
            }
            for item in respostas_temporarias
        }

        st.session_state.dados_adicionais_casos[nome_caso] = dict(dados_adicionais_temp)

        TEXTO_SEM_CONSIDERACOES = "Sem considerações específicas sobre o caso."

        if not texto_para_ia.strip():
            # Nenhuma frase relevante foi gerada (todas as respostas eram "Sim" e sem
            # observações/considerações adicionais). Não faz sentido mandar um texto
            # vazio para a IA, então usamos direto um texto padrão para o caso.
            st.session_state.casos_salvos[nome_caso] = texto_bruto
            st.session_state.consideracoes_caso[nome_caso] = consideracoes_caso
            st.session_state.identificacao_exames[nome_caso] = id_exame
            st.session_state.escolhas_casos[nome_caso] = escolhas
            st.session_state.relatorios_ia[nome_caso] = TEXTO_SEM_CONSIDERACOES
            st.session_state.docx_bytes = None

            st.success(f"{nome_caso} processado com sucesso!")
            st.rerun()
        else:
            with st.spinner("IA está formatando o relatório..."):
                try:
                    prompt = (
                        f"Deixe essas frases em um único texto coeso, não é necessário acrescentar nada, apenas o texto coeso é o suficiente. Além disso, organize as ideias apresentadas sem mudar o conteúdo."
                        f"Não mude o conteúdo, apenas deixe o texto coeso para o {nome_caso}: {texto_para_ia}"
                    )
                    response = model.generate_content(prompt)

                    st.session_state.casos_salvos[nome_caso] = texto_bruto
                    st.session_state.consideracoes_caso[nome_caso] = consideracoes_caso
                    st.session_state.identificacao_exames[nome_caso] = id_exame
                    st.session_state.escolhas_casos[nome_caso] = escolhas
                    st.session_state.relatorios_ia[nome_caso] = response.text
                    st.session_state.docx_bytes = None

                    st.success(f"{nome_caso} processado com sucesso!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao gerar relatório: {e}")

# ---------------------------------------------------------------------------
# Considerações gerais
# ---------------------------------------------------------------------------
st.markdown("---")
st.subheader("Resumo dos casos")
st.session_state.consideracoes_gerais = st.text_area(
    "Digite aqui observações que se aplicam a todos os casos:",
    value=st.session_state.consideracoes_gerais,
    height=120,
    key="consideracoes_gerais_area",
)

# ---------------------------------------------------------------------------
# Histórico da sessão
# ---------------------------------------------------------------------------
if st.session_state.relatorios_ia:
    st.markdown("---")
    st.header("Histórico da Sessão")

    casos_ordenados = sorted(st.session_state.relatorios_ia.keys(), key=extrair_numero)
    abas = st.tabs([f"{c}" for c in casos_ordenados])
    for i, nome in enumerate(casos_ordenados):
        with abas[i]:
            if st.session_state.identificacao_exames.get(nome, "").strip():
                st.markdown(f"**Identificação do Exame:** {st.session_state.identificacao_exames[nome]}")
            with st.expander("Ver texto bruto"):
                st.caption(st.session_state.casos_salvos[nome])
            if st.session_state.consideracoes_caso.get(nome, "").strip():
                st.info(st.session_state.consideracoes_caso[nome])
            st.markdown("**Relatório gerado:**")
            st.write(st.session_state.relatorios_ia[nome])

# ---------------------------------------------------------------------------
# Relatório geral (quando há pelo menos 2 casos)
# ---------------------------------------------------------------------------
if len(st.session_state.casos_salvos) >= 2:
    st.markdown("---")
    if st.button("Gerar Relatório Geral", type="primary", use_container_width=True):
        compilado = "".join([f"\n[{k}]: {v}\n" for k, v in st.session_state.casos_salvos.items()])
        texto_geral_para_ia = compilado
        if st.session_state.consideracoes_gerais.strip():
            texto_geral_para_ia += f"\n\nConsiderações gerais do avaliador: {st.session_state.consideracoes_gerais}"

        prompt_geral = (
            "Com base nos relatórios individuais abaixo, elabore um único parágrafo resumindo os achados gerais. "
            "Não mencione os números dos casos, apenas faça um resumo conciso.\n\n"
            f"Relatórios:\n{texto_geral_para_ia}"
        )
        try:
            response_geral = model.generate_content(prompt_geral)
            st.session_state.relatorio_geral_salvo = response_geral.text
            st.session_state.docx_bytes = None
            st.success("Relatório geral gerado!")
            st.rerun()
        except Exception as e:
            st.error(f"Erro ao gerar relatório geral: {e}")

    if st.session_state.relatorio_geral_salvo:
        st.markdown("### Relatório Geral ")
        st.info(st.session_state.relatorio_geral_salvo)
        st.markdown("---")
        casos_ordenados = sorted(st.session_state.relatorios_ia.keys(), key=extrair_numero)
        for nome_grupo, questoes in perguntas.items():
            st.subheader(f"Tabela de Respostas - {nome_grupo}")
            perguntas_ordenadas = list(questoes.keys())
            st.markdown(gerar_tabela_html(casos_ordenados, perguntas_ordenadas, st.session_state.escolhas_casos), unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Exportação do documento Word
# ---------------------------------------------------------------------------
if st.session_state.relatorios_ia:
    st.markdown("---")
    st.header("Exportar Documento")

    def limpar_formatacao(texto):
        return texto.replace("**", "").replace("__", "").replace("#", "")

    with st.expander("Visualizar Prévia do Documento", expanded=True):
        cab = st.session_state.dados_cabecalho
        st.markdown(f"""
        <div class='preview-box'>
            <strong>Cabeçalho</strong><br>
            <strong>Mamógrafo:</strong> {cab['mamografo_fabricante']} - {cab['mamografo_modelo']}<br>
            <strong>CNES:</strong> {cab['cnes']} &nbsp;|&nbsp; <strong>QIID:</strong> {cab['qiid']}<br>
            <strong>Tipo:</strong> {cab['tipo_mamografo']}<br>
            <strong>Instituição:</strong> {cab['instituicao']}<br>
            <strong>Cidade/Estado:</strong> {cab['cidade']} - {cab['estado']}
        </div>
        """, unsafe_allow_html=True)

        casos_ordenados = sorted(st.session_state.relatorios_ia.keys(), key=extrair_numero)
        for nome_grupo, questoes in perguntas.items():
            st.markdown(f"**Tabela de Respostas - {nome_grupo}**")
            perguntas_ordenadas = list(questoes.keys())
            st.markdown(gerar_tabela_html(casos_ordenados, perguntas_ordenadas, st.session_state.escolhas_casos), unsafe_allow_html=True)

        st.markdown("**Identificação dos Exames**")
        id_tabela = "| Caso | Identificação do Exame |\n| --- | --- |\n"
        for caso in casos_ordenados:
            id_texto = st.session_state.identificacao_exames.get(caso, "")
            id_tabela += f"| {caso} | {id_texto} |\n"
        st.markdown(id_tabela)

        st.markdown("**Dados Adicionais dos Casos**")
        adicional_tabela = "| Pergunta | " + " | ".join(casos_ordenados) + " |\n"
        adicional_tabela += "| --- | " + " | ".join(["---"] * len(casos_ordenados)) + " |\n"
        for pergunta_extra in perguntas_adicionais_texto:
            linha = [st.session_state.dados_adicionais_casos.get(caso, {}).get(pergunta_extra, "") for caso in casos_ordenados]
            adicional_tabela += f"| {pergunta_extra} | " + " | ".join(linha) + " |\n"
        st.markdown(adicional_tabela)
        st.markdown("---")

        for nome_caso in casos_ordenados:
            st.markdown(f"**{nome_caso}**")
            st.write(st.session_state.relatorios_ia[nome_caso])
            if st.session_state.consideracoes_caso.get(nome_caso, "").strip():
                st.info(st.session_state.consideracoes_caso[nome_caso])
            st.markdown("---")
        if st.session_state.relatorio_geral_salvo:
            st.markdown("**Relatório Geral **")
            st.write(st.session_state.relatorio_geral_salvo)

    def criar_docx_limpo():
        doc = Document()
        estilizar_documento(doc)

        doc.add_heading("Instrumento para a análise da qualidade da mamografia", level=0)
        doc.add_paragraph()

        cab = st.session_state.dados_cabecalho
        p = doc.add_paragraph()
        p.add_run("Mamógrafo (fabricante e modelo): ").bold = True
        p.add_run(f"{cab['mamografo_fabricante']} - {cab['mamografo_modelo']}")

        p = doc.add_paragraph()
        p.add_run("CNES: ").bold = True
        p.add_run(cab["cnes"])
        p.add_run("     QIID: ").bold = True
        p.add_run(cab["qiid"])

        p = doc.add_paragraph()
        p.add_run("Tipo de mamógrafo: ").bold = True
        opcoes_tipo = ["Convencional", "Digital CR", "Digital DR", "DR retrofit"]
        for opcao in opcoes_tipo:
            marcado = "X" if cab["tipo_mamografo"] == opcao else "-"
            p.add_run(f"  [{marcado}] {opcao}  ")

        doc.add_paragraph()
        p = doc.add_paragraph()
        p.add_run("Instituição: ").bold = True
        p.add_run(cab["instituicao"])

        p = doc.add_paragraph()
        p.add_run("Cidade: ").bold = True
        p.add_run(cab["cidade"])
        p.add_run("     Estado: ").bold = True
        p.add_run(cab["estado"])
        doc.add_paragraph()

        # Tabelas de respostas - uma por grupo
        casos_ord = sorted(st.session_state.relatorios_ia.keys(), key=extrair_numero)
        num_casos = len(casos_ord)

        for nome_grupo, questoes in perguntas.items():
            doc.add_heading(f"Tabela de Respostas - {nome_grupo}", level=1)
            perguntas_ord = list(questoes.keys())
            total_colunas = 1 + num_casos * 2
            tabela = doc.add_table(rows=2 + len(perguntas_ord), cols=total_colunas)
            tabela.style = "Table Grid"

            tabela.cell(0, 0).merge(tabela.cell(1, 0))
            formatar_celula(tabela.cell(0, 0), "Pergunta", cor_fundo=COR_PRIMARIA, cor_texto="FFFFFF", negrito=True)

            for idx, caso in enumerate(casos_ord):
                col_inicio = 1 + idx * 2
                col_fim = col_inicio + 1
                tabela.cell(0, col_inicio).merge(tabela.cell(0, col_fim))
                formatar_celula(tabela.cell(0, col_inicio), caso, cor_fundo=COR_SECUNDARIA, cor_texto="FFFFFF", negrito=True)

            for idx in range(num_casos):
                col_sim = 1 + idx * 2
                col_nao = col_sim + 1
                formatar_celula(tabela.cell(1, col_sim), "Sim", cor_fundo=COR_PRIMARIA, cor_texto="FFFFFF", negrito=True)
                formatar_celula(tabela.cell(1, col_nao), "Não", cor_fundo=COR_PRIMARIA, cor_texto="FFFFFF", negrito=True)

            for i, pergunta in enumerate(perguntas_ord):
                linha_atual = i + 2
                cor_linha = COR_ZEBRA if i % 2 == 1 else None
                formatar_celula(tabela.cell(linha_atual, 0), pergunta, cor_fundo=cor_linha, centralizar=False)
                for j, caso in enumerate(casos_ord):
                    item = st.session_state.escolhas_casos.get(caso, {}).get(pergunta, {})
                    resposta = item.get("resposta", "-") if isinstance(item, dict) else item
                    col_sim = 1 + j * 2
                    col_nao = col_sim + 1
                    formatar_celula(tabela.cell(linha_atual, col_sim), "X" if resposta == "Sim" else "", cor_fundo=cor_linha, negrito=True)
                    formatar_celula(tabela.cell(linha_atual, col_nao), "X" if resposta == "Não" else "", cor_fundo=cor_linha, negrito=True)

            doc.add_paragraph()

        doc.add_page_break()

        # Identificação dos exames
        doc.add_heading("Identificação dos Exames", level=2)
        mini_tabela = doc.add_table(rows=2, cols=num_casos)
        mini_tabela.style = "Table Grid"
        for j, caso in enumerate(casos_ord):
            formatar_celula(mini_tabela.rows[0].cells[j], caso, cor_fundo=COR_PRIMARIA, cor_texto="FFFFFF", negrito=True)
        for j, caso in enumerate(casos_ord):
            formatar_celula(mini_tabela.rows[1].cells[j], st.session_state.identificacao_exames.get(caso, ""))
        doc.add_paragraph()

        # Dados adicionais dos casos (não entram no texto da IA)
        doc.add_heading("Dados Adicionais dos Casos", level=2)
        tabela_adicional = doc.add_table(rows=1 + len(perguntas_adicionais_texto), cols=1 + num_casos)
        tabela_adicional.style = "Table Grid"
        formatar_celula(tabela_adicional.cell(0, 0), "Pergunta", cor_fundo=COR_PRIMARIA, cor_texto="FFFFFF", negrito=True)
        for j, caso in enumerate(casos_ord):
            formatar_celula(tabela_adicional.cell(0, j + 1), caso, cor_fundo=COR_PRIMARIA, cor_texto="FFFFFF", negrito=True)
        for i, pergunta_extra in enumerate(perguntas_adicionais_texto):
            linha_atual = i + 1
            cor_linha = COR_ZEBRA if i % 2 == 1 else None
            formatar_celula(tabela_adicional.cell(linha_atual, 0), pergunta_extra, cor_fundo=cor_linha, centralizar=False)
            for j, caso in enumerate(casos_ord):
                valor = st.session_state.dados_adicionais_casos.get(caso, {}).get(pergunta_extra, "")
                formatar_celula(tabela_adicional.cell(linha_atual, j + 1), valor, cor_fundo=cor_linha)
        doc.add_paragraph()

        # Anexo
        doc.add_heading("Anexo ao instrumento para análise da qualidade da mamografia", level=0)
        p = doc.add_paragraph()
        p.add_run("Serviço: ").bold = True
        p.add_run(cab["servico"])
        p = doc.add_paragraph()
        p.add_run("CNES: ").bold = True
        p.add_run(cab["cnes"])
        p = doc.add_paragraph()
        p.add_run("QIID: ").bold = True
        p.add_run(cab["qiid"])
        doc.add_paragraph()

        # Considerações específicas
        doc.add_heading("Considerações Específicas", level=0)
        for nome_caso in casos_ord:
            doc.add_heading(nome_caso, level=1)
            texto_ia = st.session_state.relatorios_ia[nome_caso]
            for linha in texto_ia.strip().split("\n"):
                if linha.strip():
                    doc.add_paragraph(limpar_formatacao(linha))
            if st.session_state.consideracoes_caso.get(nome_caso, "").strip():
                doc.add_heading("Considerações Adicionais", level=2)
                doc.add_paragraph(limpar_formatacao(st.session_state.consideracoes_caso[nome_caso]))
            doc.add_paragraph("-" * 30)

        # Recomendações
        # Estrutura: {pergunta: {sub_opcao: texto}}.
        # Perguntas sem sub_opcoes usam a chave "_default".
        # Perguntas com sub_opcoes podem ter um texto por sub-opção; se uma
        # sub-opção não tiver entrada própria, cai no "_default" da pergunta
        recomendacoes = {
            "Utiliza corretamente o Léxico BI-RADS® ou SISMAMA":{
                "_default":
                    "É recomendado aos médicos do serviço a realização do curso de reciclagem em diagnóstico mamário “Atualização em BI-RADS”, oferecido pelo Colégio Brasileiro de Radiologia. O curso é gratuito para os médicos dos serviços que participam do Programa de Qualidade em Mamografia do INCA. O detalhamento do processo para realização do curso indicado será encaminhado em um e-mail à parte, que tratará exclusivamente desse assunto."
            },
        }

        # ---------------------------------------------------------------
        # Recomendações "por grupo": diferente das recomendações acima
        # (que são ligadas a UMA pergunta específica), estas disparam se
        # QUALQUER pergunta de um grupo (menos as que você excluir) tiver
        # a resposta indicada em "resposta_gatilho". A recomendação
        # aparece só UMA VEZ no documento, mesmo que várias perguntas do
        # grupo tenham disparado para o mesmo caso ou para casos diferentes.
        #
        # Para CRIAR uma nova recomendação de grupo, copie um bloco abaixo
        # e ajuste "grupo", "resposta_gatilho", "perguntas_excluidas" (ou
        # "perguntas_incluidas") e "texto".
        # Para REMOVER, basta apagar o bloco correspondente da lista.
        #
        # - "perguntas_excluidas": lista de perguntas do grupo que NÃO
        #   devem contar para disparar essa recomendação (as demais do
        #   grupo contam automaticamente).
        # - "perguntas_incluidas": se preenchida, usa EXATAMENTE essa
        #   lista de perguntas (ignora "perguntas_excluidas"). Útil se
        #   você quiser escolher a dedo quais perguntas participam, em
        #   vez de excluir só uma ou duas.
        # ---------------------------------------------------------------
        recomendacoes_por_grupo = [
            # Nenhuma recomendação de grupo ativa no momento (o grupo de
            # posicionamento, que usava essa estrutura, foi removido).
            # Para adicionar uma nova, copie o formato do bloco de exemplo
            # nos comentários acima.
        ]

        def resposta_do_caso(caso, pergunta):
            item = st.session_state.escolhas_casos.get(caso, {}).get(pergunta, {})
            return item.get("resposta", "") if isinstance(item, dict) else item

        def sub_opcao_do_caso(caso, pergunta):
            item = st.session_state.escolhas_casos.get(caso, {}).get(pergunta, {})
            return item.get("sub_opcao") if isinstance(item, dict) else None

        def obter_gatilho(pergunta):
            # Busca o gatilho configurado em `perguntas` (padrão: "Não").
            for questoes in perguntas.values():
                if pergunta in questoes:
                    return questoes[pergunta].get("gatilho_sub_opcoes", "Não")
            return "Não"

        def perguntas_da_config_grupo(config):
            questoes_grupo = perguntas.get(config["grupo"], {})
            incluidas = config.get("perguntas_incluidas") or []
            if incluidas:
                return incluidas
            excluidas = set(config.get("perguntas_excluidas") or [])
            return [p for p in questoes_grupo if p not in excluidas]

        def grupo_disparado(caso, config):
            resposta_gatilho = config.get("resposta_gatilho", "Não")
            return any(
                resposta_do_caso(caso, pergunta) == resposta_gatilho
                for pergunta in perguntas_da_config_grupo(config)
            )

        tem_recomendacao = any(
            resposta_do_caso(caso, pergunta) == obter_gatilho(pergunta)
            for caso in casos_ord
            for pergunta in recomendacoes
        ) or any(
            grupo_disparado(caso, config)
            for caso in casos_ord
            for config in recomendacoes_por_grupo
        )
        if tem_recomendacao:
            doc.add_heading("Recomendações", level=0)
            textos_inseridos = set()
            for caso in casos_ord:
                for pergunta, textos_por_sub in recomendacoes.items():
                    if resposta_do_caso(caso, pergunta) == obter_gatilho(pergunta):
                        subs = sub_opcao_do_caso(caso, pergunta)
                        if not subs:
                            # Pergunta sem sub_opcoes (ou nenhuma sub-opção marcada): usa o texto padrão.
                            textos = [textos_por_sub.get("_default")]
                        else:
                            # Uma ou mais sub-opções marcadas: junta o texto de cada uma.
                            textos = [
                                textos_por_sub.get(sub) or textos_por_sub.get("_default")
                                for sub in subs
                            ]
                        for texto in textos:
                            if texto and texto not in textos_inseridos:
                                textos_inseridos.add(texto)
                                p = doc.add_paragraph(style="List Bullet")
                                p.add_run(texto)

            for caso in casos_ord:
                for config in recomendacoes_por_grupo:
                    if grupo_disparado(caso, config):
                        texto = config.get("texto")
                        if texto and texto not in textos_inseridos:
                            textos_inseridos.add(texto)
                            p = doc.add_paragraph(style="List Bullet")
                            p.add_run(texto)

        if st.session_state.relatorio_geral_salvo:
            doc.add_paragraph()
            p = doc.add_paragraph()
            p.add_run("Resumo dos casos:").bold = True
            for linha in st.session_state.relatorio_geral_salvo.strip().split("\n"):
                if linha.strip():
                    doc.add_paragraph(limpar_formatacao(linha))

        output = BytesIO()
        doc.save(output)
        output.seek(0)
        return output

    if st.button("Gerar Documento Word", use_container_width=True):
        with st.spinner("Montando o documento..."):
            st.session_state.docx_bytes = criar_docx_limpo().getvalue()
        st.success("Documento gerado! Use o botão abaixo para baixar.")

    if st.session_state.get("docx_bytes"):
        nome_servico_arquivo = sanitizar_nome_arquivo(st.session_state.dados_cabecalho.get("servico", ""))
        st.download_button(
            label="Baixar Documento Final (.docx)",
            data=st.session_state.docx_bytes,
            file_name=f"relatório_{nome_servico_arquivo}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

# ---------------------------------------------------------------------------
# Botão de reset
# ---------------------------------------------------------------------------
st.markdown("---")
if st.button("Limpar todos os dados da sessão"):
    for chave in [
        "casos_salvos",
        "relatorios_ia",
        "relatorio_geral_salvo",
        "consideracoes_caso",
        "consideracoes_gerais",
        "escolhas_casos",
        "dados_cabecalho",
        "identificacao_exames",
        "dados_adicionais_casos",
        "docx_bytes",
    ]:
        if chave in st.session_state:
            del st.session_state[chave]
    st.rerun()

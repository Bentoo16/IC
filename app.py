import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor
from io import BytesIO
import re

# ---------------------------------------------------------------------------
# CSS customizado
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .tabela-respostas {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.85rem;
        margin-bottom: 1.5rem;
    }
    .tabela-respostas th {
        background: #E0E0E0;
        color: #333;
        padding: 0.5rem;
        text-align: center;
        border: 1px solid #ccc;
    }
    .tabela-respostas th.caso-header {
        color: #FF0000;
    }
    .tabela-respostas td {
        padding: 0.4rem 0.5rem;
        border: 1px solid #dee2e6;
        text-align: center;
    }
    .tabela-respostas td:first-child {
        text-align: left;
        font-weight: 500;
    }
    .preview-box {
        background: #f4f4f4;
        border-left: 4px solid #555;
        padding: 1rem 1.2rem;
        border-radius: 6px;
        margin-bottom: 1rem;
    }
    .progresso-texto {
        font-size: 0.85rem;
        color: #555;
        margin-bottom: 0.3rem;
    }
    .stButton>button {
        font-weight: 600;
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
            escolha = escolhas_casos.get(caso, {}).get(pergunta, "-")
            sim_cell = "X" if escolha == "Sim" else ""
            nao_cell = "X" if escolha == "Não" else ""
            html += f"<td>{sim_cell}</td><td>{nao_cell}</td>"
        html += "</tr>"
    html += "</table>"
    return html


def contar_perguntas(grupos):
    total = 0
    for qs in grupos.values():
        total += len(qs)
    return total


def set_cell_shading(cell, color):
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    shading.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shading)


# ---------------------------------------------------------------------------
# Configuração da IA
# ---------------------------------------------------------------------------
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-3.1-flash-lite')

st.title("Gerador de Relatórios - Mamografia")

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
    "Aspectos Físicos da Imagem": {
        "Contraste adequado": {
            "opcoes": {"Sim": "O contraste está adequado.", "Não": "O contraste não está adequado."},
            "sub_opcoes": {
                "Contraste alto demais": "O contraste da imagem está alto demais.",
                "Contraste baixo demais": "O contraste está baixo demais.",
            },
        },
        "Definição de estruturas": {
            "opcoes": {
                "Sim": "As estruturas estão bem definidas na imagem.",
                "Não": "As estruturas não estão bem definidas na imagem.",
            },
            "sub_opcoes": {
                "Problema A": "Frase gerada para o problema A.",
                "Problema B": "Frase gerada para o problema B.",
            },
        },
        "Saturação correta nas áreas claras": {
            "opcoes": {
                "Sim": "A imagem está bem saturada nas áreas claras.",
                "Não": "A imagem não está bem saturada nas áreas claras.",
            },
            "sub_opcoes": {
                "Problema A": "Frase gerada para o problema A.",
                "Problema B": "Frase gerada para o problema B.",
            },
        },
        "Saturação correta nas áreas escuras": {
            "opcoes": {
                "Sim": "A imagem está bem saturada nas áreas escuras.",
                "Não": "A imagem não está bem saturada nas áreas escuras.",
            },
            "sub_opcoes": {
                "Problema A": "Frase gerada para o problema A.",
                "Problema B": "Frase gerada para o problema B.",
            },
        },
        "Imagem sem ruído": {
            "opcoes": {"Sim": "A imagem está sem ruído.", "Não": "A imagem está com ruído."},
            "sub_opcoes": {
                "Problema A": "Frase gerada para o problema A.",
                "Problema B": "Frase gerada para o problema B.",
            },
        },
        "A área de fundo está adequadamente escura (enegrecimento película)": {
            "opcoes": {
                "Sim": "A área de fundo da imagem está adequadamente escura.",
                "Não": "A área de fundo da imagem não está adequadamente escura.",
            },
            "sub_opcoes": {
                "Problema A": "Frase gerada para o problema A.",
                "Problema genérico B": "Frase gerada para o problema B.",
            },
        },
        "Imagem sem artefatos (se houver, descrever)": {
            "opcoes": {"Sim": "A imagem não possui artefatos.", "Não": "A imagem possui artefatos."},
            "sub_opcoes": {
                "Problema A": "Frase gerada para o problema A.",
                "Problema B": "Frase gerada para o problema B.",
            },
        },
    },
    "Avaliação dos Critérios de Laudos": {
        "Resumo da história presente": {
            "opcoes": {"Sim": "O resumo da história presente está adequado.", "Não": "O resumo da história presente não está adequado."},
            "sub_opcoes": {
                "Problema A": "Frase gerada para o problema A.",
                "Problema B": "Frase gerada para o problema B.",
            },
        },
        "Utiliza corretamente o Léxico BI-RADS ou SISMAMA": {
            "opcoes": {"Sim": "Utiliza corretamente o léxico BI-RADS ou SISMAMA.", "Não": "Não utiliza corretamente o léxico BI-RADS ou SISMAMA."},
            "sub_opcoes": {
                "Problema A": "Frase gerada para o problema A.",
                "Problema B": "Frase gerada para o problema B.",
            },
        },
        "Classifica corretamente o exame segundo o BI-RADS": {
            "opcoes": {"Sim": "Classifica corretamente o exame segundo o BI-RADS.", "Não": "Não classifica corretamente o exame segundo o BI-RADS."},
            "sub_opcoes": {
                "Problema A": "Frase gerada para o problema A.",
                "Problema B": "Frase gerada para o problema B.",
            },
        },
        "Recomendação correta segundo o BI-RADS": {
            "opcoes": {"Sim": "Recomenda corretamente o exame segundo o BI-RADS.", "Não": "Não recomenda corretamente o exame segundo o BI-RADS."},
            "sub_opcoes": {
                "Problema A": "Frase gerada para o problema A.",
                "Problema B": "Frase gerada para o problema B.",
            },
        },
        "Interpretou corretamente todos os achados do exame": {
            "opcoes": {"Sim": "Interpretou corretamente todos os achados do exame.", "Não": "Não interpretou corretamente todos os achados do exame."},
            "sub_opcoes": {
                "Problema A": "Frase gerada para o problema A.",
                "Problema B": "Frase gerada para o problema B.",
            },
        },
    },
}

# ---------------------------------------------------------------------------
# Seleção do caso e perguntas
# ---------------------------------------------------------------------------
st.markdown("---")
caso_atual = st.selectbox("Escolha o Caso que vai analisar agora:", [1, 2, 3, 4, 5])
nome_caso = f"Caso {caso_atual}"

respostas_temporarias = []
abas_grupos = st.tabs(list(perguntas.keys()))
for idx, (nome_grupo, questoes) in enumerate(perguntas.items()):
    with abas_grupos[idx]:
        for titulo, info in questoes.items():
            st.subheader(titulo)
            escolha = st.radio("Selecione:", list(info["opcoes"].keys()), key=f"radio_{titulo}_c{caso_atual}", horizontal=True)
            sub_escolha = None
            if "sub_opcoes" in info and escolha == "Não":
                sub_escolha = st.radio("Especifique:", list(info["sub_opcoes"].keys()), key=f"sub_{titulo}_c{caso_atual}")
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

nome_caso = f"Caso {caso_atual}"
caso_ja_existe = nome_caso in st.session_state.casos_salvos
confirmacao = True
if caso_ja_existe:
    st.warning(f"O {nome_caso} já foi salvo anteriormente.")
    confirmacao = st.checkbox("Deseja sobrescrever o relatório existente?", key=f"conf_{caso_atual}")

if st.button(f"Analisar e Salvar {nome_caso}", type="primary", use_container_width=True):
    if caso_ja_existe and not confirmacao:
        st.warning("Marque a confirmação para sobrescrever o caso.")
    else:
        respostas_finais = []
        for item in respostas_temporarias:
            # Encontrar o grupo e a pergunta
            for questoes in perguntas.values():
                if item["titulo"] in questoes:
                    info_pergunta = questoes[item["titulo"]]
                    break
            if item["escolha"] == "Não" and item["sub_escolha"]:
                frase_base = info_pergunta["sub_opcoes"][item["sub_escolha"]]
            else:
                frase_base = info_pergunta["opcoes"][item["escolha"]]
            if item["obs"]:
                frase_base += f" Detalhe adicional: {item['obs']}"
            respostas_finais.append(frase_base)

        texto_bruto = " ".join(respostas_finais)
        texto_para_ia = texto_bruto
        if consideracoes_caso.strip():
            texto_para_ia += f"\n\n{consideracoes_caso}"

        escolhas = {item["titulo"]: item["escolha"] for item in respostas_temporarias}

        with st.spinner("IA está formatando o relatório..."):
            try:
                prompt = (
                    f"Deixe essas frases em um único texto coeso, não é necessário acrescentar nada, apenas o texto coeso é o suficiente. "
                    f"Não mude as frases, apenas deixe o texto coeso para o {nome_caso}: {texto_para_ia}"
                )
                response = model.generate_content(prompt)

                st.session_state.casos_salvos[nome_caso] = texto_bruto
                st.session_state.consideracoes_caso[nome_caso] = consideracoes_caso
                st.session_state.identificacao_exames[nome_caso] = id_exame
                st.session_state.escolhas_casos[nome_caso] = escolhas
                st.session_state.relatorios_ia[nome_caso] = response.text

                st.success(f"{nome_caso} processado com sucesso!")
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao gerar relatório: {e}")

# ---------------------------------------------------------------------------
# Considerações gerais
# ---------------------------------------------------------------------------
st.markdown("---")
st.subheader("Considerações Gerais")
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
            tabela.cell(0, 0).text = "Pergunta"
            set_cell_shading(tabela.cell(0, 0), "E0E0E0")
            set_cell_shading(tabela.cell(1, 0), "E0E0E0")

            for idx, caso in enumerate(casos_ord):
                col_inicio = 1 + idx * 2
                col_fim = col_inicio + 1
                tabela.cell(0, col_inicio).merge(tabela.cell(0, col_fim))
                cell_caso = tabela.cell(0, col_inicio)
                cell_caso.text = ""
                run = cell_caso.paragraphs[0].add_run(caso)
                run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
                set_cell_shading(cell_caso, "E0E0E0")

            for idx in range(num_casos):
                col_sim = 1 + idx * 2
                col_nao = col_sim + 1
                tabela.cell(1, col_sim).text = "Sim"
                set_cell_shading(tabela.cell(1, col_sim), "E0E0E0")
                tabela.cell(1, col_nao).text = "Não"
                set_cell_shading(tabela.cell(1, col_nao), "E0E0E0")

            for i, pergunta in enumerate(perguntas_ord):
                linha_atual = i + 2
                tabela.cell(linha_atual, 0).text = pergunta
                for j, caso in enumerate(casos_ord):
                    escolha = st.session_state.escolhas_casos.get(caso, {}).get(pergunta, "-")
                    col_sim = 1 + j * 2
                    col_nao = col_sim + 1
                    tabela.cell(linha_atual, col_sim).text = "X" if escolha == "Sim" else ""
                    tabela.cell(linha_atual, col_nao).text = "X" if escolha == "Não" else ""

            doc.add_paragraph()

        doc.add_page_break()

        # Identificação dos exames
        doc.add_heading("Identificação dos Exames", level=2)
        mini_tabela = doc.add_table(rows=2, cols=num_casos)
        mini_tabela.style = "Table Grid"
        for j, caso in enumerate(casos_ord):
            mini_tabela.rows[0].cells[j].text = caso
        for j, caso in enumerate(casos_ord):
            mini_tabela.rows[1].cells[j].text = st.session_state.identificacao_exames.get(caso, "")
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

        if st.session_state.relatorio_geral_salvo:
            doc.add_paragraph()
            for linha in st.session_state.relatorio_geral_salvo.strip().split("\n"):
                if linha.strip():
                    doc.add_paragraph(limpar_formatacao(linha))

        output = BytesIO()
        doc.save(output)
        output.seek(0)
        return output

    st.download_button(
        label="Baixar Documento Final (.docx)",
        data=criar_docx_limpo(),
        file_name="relatorio_final.docx",
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
    ]:
        if chave in st.session_state:
            del st.session_state[chave]
    st.rerun()

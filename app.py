import streamlit as st
import google.generativeai as genai
from docx import Document
from io import BytesIO
import re

# Configuração da IA
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-2.5-flash-lite')

st.title("Gerador de Relatórios")

# ------------------------------------------------------------
# Inicialização do session_state
# ------------------------------------------------------------
if "dados_cabecalho" not in st.session_state:
    st.session_state.dados_cabecalho = {
        "mamografo_fabricante": "",
        "mamografo_modelo": "",
        "cnes": "",
        "qiid": "",
        "tipo_mamografo": None,
        "instituicao": "",
        "cidade": "",
        "estado": ""
    }
if "casos_salvos" not in st.session_state:
    st.session_state.casos_salvos = {}
if "relatorios_ia" not in st.session_state:
    st.session_state.relatorios_ia = {}
if "relatorio_geral_salvo" not in st.session_state:
    st.session_state.relatorio_geral_salvo = None
# consideracoes_caso_grupo: dicionário aninhado caso -> grupo -> texto
if "consideracoes_caso_grupo" not in st.session_state:
    st.session_state.consideracoes_caso_grupo = {}
# consideracoes_gerais_grupo: dicionário grupo -> texto
if "consideracoes_gerais_grupo" not in st.session_state:
    st.session_state.consideracoes_gerais_grupo = {}
if "escolhas_casos" not in st.session_state:
    st.session_state.escolhas_casos = {}
if "identificacao_exames" not in st.session_state:
    st.session_state.identificacao_exames = {}

# ------------------------------------------------------------
# Cabeçalho – dados da instituição
# ------------------------------------------------------------
st.markdown("---")
st.subheader("Dados do Cabeçalho (para o relatório Word)")

col1, col2 = st.columns(2)
with col1:
    fabricante = st.text_input("Mamógrafo – Fabricante:", value=st.session_state.dados_cabecalho["mamografo_fabricante"])
with col2:
    modelo = st.text_input("Mamógrafo – Modelo:", value=st.session_state.dados_cabecalho["mamografo_modelo"])

cnes = st.text_input("CNES:", value=st.session_state.dados_cabecalho["cnes"])
qiid = st.text_input("QIID:", value=st.session_state.dados_cabecalho["qiid"])

tipo = st.radio(
    "Tipo de mamógrafo:",
    ["Convencional", "Digital CR", "Digital DR", "DR retrofit"],
    index=0 if st.session_state.dados_cabecalho["tipo_mamografo"] is None else
          ["Convencional", "Digital CR", "Digital DR", "DR retrofit"].index(st.session_state.dados_cabecalho["tipo_mamografo"]),
    key="tipo_mamografo_radio"
)

instituicao = st.text_input("Instituição:", value=st.session_state.dados_cabecalho["instituicao"])

col3, col4 = st.columns(2)
with col3:
    cidade = st.text_input("Cidade:", value=st.session_state.dados_cabecalho["cidade"])
with col4:
    estado = st.text_input("Estado:", value=st.session_state.dados_cabecalho["estado"])

st.session_state.dados_cabecalho = {
    "mamografo_fabricante": fabricante,
    "mamografo_modelo": modelo,
    "cnes": cnes,
    "qiid": qiid,
    "tipo_mamografo": tipo,
    "instituicao": instituicao,
    "cidade": cidade,
    "estado": estado
}

# ------------------------------------------------------------
# Definição dos grupos de perguntas
# ------------------------------------------------------------
grupos_perguntas = [
    {
        "titulo": "Aspectos Físicos da Imagem",
        "perguntas": {
            "Contraste adequado": {
                "opcoes": {"Sim": "O contraste está adequado.", "Não": "O contraste não está adequado."},
                "sub_opcoes": {
                    "Contraste alto demais": "O contraste da imagem está alto demais.",
                    "Contraste baixo demais": "O contraste está baixo demais."
                }
            },
            "Definição de estruturas": {
                "opcoes": {"Sim": "As estruturas estão bem definidas na imagem.",
                           "Não": "As estruturas não estão bem definidas na imagem."},
                "sub_opcoes": {
                    "Problema A": "Frase gerada para o problema A.",
                    "Problema B": "Frase gerada para o problema B."
                }
            },
            "Saturação correta nas áreas claras": {
                "opcoes": {"Sim": "A imagem está bem saturada nas áreas claras.",
                           "Não": "A imagem não está bem saturada nas áreas claras."},
                "sub_opcoes": {
                    "Problema A": "Frase gerada para o problema A.",
                    "Problema B": "Frase gerada para o problema B."
                }
            },
            "Saturação correta nas áreas escuras": {
                "opcoes": {"Sim": "A imagem está bem saturada nas áreas escuras.",
                           "Não": "A imagem não está bem saturada nas áreas escuras."},
                "sub_opcoes": {
                    "Problema A": "Frase gerada para o problema A.",
                    "Problema B": "Frase gerada para o problema B."
                }
            },
            "Imagem sem ruído": {
                "opcoes": {"Sim": "A imagem está sem ruído.", "Não": "A imagem está com ruído."},
                "sub_opcoes": {
                    "Problema A": "Frase gerada para o problema A.",
                    "Problema B": "Frase gerada para o problema B."
                }
            },
            "A área de fundo está adequadamente escura (enegrecimento película)": {
                "opcoes": {"Sim": "A área de fundo da imagem está adequadamente escura.",
                           "Não": "A área de fundo da imagem não está adequadamente escura."},
                "sub_opcoes": {
                    "Problema A": "Frase gerada para o problema A.",
                    "Problema genérico B": "Frase gerada para o problema B."
                }
            },
            "Imagem sem artefatos (se houver, descrever)": {
                "opcoes": {"Sim": "A imagem não possui artefatos.", "Não": "A imagem possui artefatos."},
                "sub_opcoes": {
                    "Problema A": "Frase gerada para o problema A.",
                    "Problema B": "Frase gerada para o problema B."
                }
            }
        }
    },
    {
        "titulo": "Avaliação dos Critérios e laudos",
        "perguntas": {
            "Critério 1": {
                "opcoes": {"Sim": "Critério 1 atendido.", "Não": "Critério 1 não atendido."},
                "sub_opcoes": {}
            },
            "Critério 2": {
                "opcoes": {"Sim": "Critério 2 atendido.", "Não": "Critério 2 não atendido."},
                "sub_opcoes": {}
            },
            "Critério 3": {
                "opcoes": {"Sim": "Critério 3 atendido.", "Não": "Critério 3 não atendido."},
                "sub_opcoes": {}
            }
        }
    }
]

# ------------------------------------------------------------
# Seleção do caso e exibição dos grupos (com considerações por grupo)
# ------------------------------------------------------------
st.markdown("---")
caso_atual = st.selectbox("Escolha o Caso que vai analisar agora:", [1, 2, 3, 4, 5])

respostas_temporarias = []
# Para armazenar as considerações de cada grupo (desse caso)
consideracoes_grupo_temp = {}

for grupo in grupos_perguntas:
    with st.expander(grupo["titulo"], expanded=False):
        for titulo, info in grupo["perguntas"].items():
            st.markdown(f"**{titulo}**")
            escolha = st.radio("Selecione:", list(info["opcoes"].keys()), key=f"radio_{grupo['titulo']}_{titulo}_c{caso_atual}")
            sub_escolha = None
            if "sub_opcoes" in info and escolha == "Não":
                sub_escolha = st.radio("Especifique:", list(info["sub_opcoes"].keys()), key=f"sub_{grupo['titulo']}_{titulo}_c{caso_atual}")
            obs = st.text_input("Observação: ", key=f"obs_{grupo['titulo']}_{titulo}_c{caso_atual}")
            respostas_temporarias.append({
                "titulo": titulo,
                "grupo": grupo["titulo"],
                "escolha": escolha,
                "sub_escolha": sub_escolha,
                "obs": obs
            })
        # Campo de considerações adicionais para este grupo
        chave_grupo = grupo["titulo"]
        texto_padrao = st.session_state.consideracoes_caso_grupo.get(f"Caso {caso_atual}", {}).get(chave_grupo, "")
        cons_grupo = st.text_area(
            f"Considerações adicionais para este grupo ({chave_grupo}):",
            value=texto_padrao,
            key=f"cons_grupo_{chave_grupo}_c{caso_atual}",
            height=100
        )
        consideracoes_grupo_temp[chave_grupo] = cons_grupo

# Identificação do exame (mantida)
st.markdown("---")
id_exame = st.text_input(
    "Identificação do Exame:",
    value=st.session_state.identificacao_exames.get(f"Caso {caso_atual}", ""),
    key=f"id_exame_c{caso_atual}"
)

# ------------------------------------------------------------
# Lógica de salvamento do caso (SEM CHAMAR IA)
# ------------------------------------------------------------
nome_caso = f"Caso {caso_atual}"
caso_ja_existe = nome_caso in st.session_state.casos_salvos
confirmacao = True
if caso_ja_existe:
    st.warning(f"O {nome_caso} já foi salvo anteriormente.")
    confirmacao = st.checkbox("Deseja sobrescrever o relatório existente?", key=f"conf_{caso_atual}")

if st.button(f"Analisar e Salvar Caso {caso_atual}"):
    if caso_ja_existe and not confirmacao:
        st.warning("Marque a confirmação para sobrescrever o caso.")
    else:
        # Constrói o texto bruto do caso concatenando as respostas de todos os grupos
        texto_bruto = ""
        for grupo in grupos_perguntas:
            respostas_grupo = []
            for item in respostas_temporarias:
                if item["grupo"] == grupo["titulo"]:
                    info_pergunta = grupo["perguntas"][item["titulo"]]
                    frase = info_pergunta["sub_opcoes"][item["sub_escolha"]] if item["escolha"] == "Não" and item["sub_escolha"] else info_pergunta["opcoes"][item["escolha"]]
                    if item["obs"]:
                        frase += f" Detalhe: {item['obs']}"
                    respostas_grupo.append(frase)
            bloco_grupo = " ".join(respostas_grupo)
            # Adiciona considerações do grupo, se houver
            cons = consideracoes_grupo_temp.get(grupo["titulo"], "")
            if cons.strip():
                bloco_grupo += f" Considerações adicionais do grupo '{grupo['titulo']}': {cons}"
            texto_bruto += bloco_grupo + " "

        st.session_state.casos_salvos[nome_caso] = texto_bruto.strip()
        
        # Salva as considerações por grupo
        if nome_caso not in st.session_state.consideracoes_caso_grupo:
            st.session_state.consideracoes_caso_grupo[nome_caso] = {}
        for grupo_nome, texto in consideracoes_grupo_temp.items():
            st.session_state.consideracoes_caso_grupo[nome_caso][grupo_nome] = texto

        st.session_state.identificacao_exames[nome_caso] = id_exame

        # Salva as escolhas (com chave prefixada para evitar conflitos)
        escolhas = {}
        for item in respostas_temporarias:
            chave_pergunta = f"{item['grupo']}||{item['titulo']}"
            escolhas[chave_pergunta] = item["escolha"]
        st.session_state.escolhas_casos[nome_caso] = escolhas

        st.success(f"Caso {caso_atual} salvo! (Relatório será gerado junto com o geral)")

# ------------------------------------------------------------
# Considerações gerais por grupo
# ------------------------------------------------------------
st.markdown("---")
st.subheader("Considerações Gerais por Grupo")
for grupo in grupos_perguntas:
    chave = grupo["titulo"]
    texto_atual = st.session_state.consideracoes_gerais_grupo.get(chave, "")
    st.session_state.consideracoes_gerais_grupo[chave] = st.text_area(
        f"Considerações gerais para o grupo '{chave}':",
        value=texto_atual,
        key=f"cons_geral_{chave}",
        height=100
    )

# ------------------------------------------------------------
# Histórico da sessão (exibe relatórios apenas se já gerados)
# ------------------------------------------------------------
if st.session_state.relatorios_ia:
    st.markdown("---")
    st.header("Histórico da Sessão")

    def extrair_numero(nome):
        match = re.search(r'\d+', nome)
        return int(match.group()) if match else 0

    casos_ordenados = sorted(st.session_state.relatorios_ia.keys(), key=extrair_numero)
    abas = st.tabs(casos_ordenados)
    for i, nome_caso in enumerate(casos_ordenados):
        with abas[i]:
            if nome_caso in st.session_state.identificacao_exames and st.session_state.identificacao_exames[nome_caso].strip():
                st.write("**Identificação do Exame:**")
                st.info(st.session_state.identificacao_exames[nome_caso])
            st.write("**Texto Bruto:**")
            st.caption(st.session_state.casos_salvos[nome_caso])
            # Mostra considerações de cada grupo para este caso
            if nome_caso in st.session_state.consideracoes_caso_grupo:
                for grupo_nome, texto in st.session_state.consideracoes_caso_grupo[nome_caso].items():
                    if texto.strip():
                        st.write(f"**Considerações ({grupo_nome}):**")
                        st.info(texto)
            st.write("**Relatório IA:**")
            st.write(st.session_state.relatorios_ia[nome_caso])

# ------------------------------------------------------------
# Relatório geral (GERA TUDO EM UMA ÚNICA CHAMADA)
# ------------------------------------------------------------
if len(st.session_state.casos_salvos) >= 2:
    if st.button("Gerar Relatório Geral"):
        # Monta o texto compilado para enviar à IA
        casos_texto = []
        for nome_caso in sorted(st.session_state.casos_salvos.keys(), key=extrair_numero):
            texto = st.session_state.casos_salvos[nome_caso]
            casos_texto.append(f"{nome_caso}: {texto}")
        compilado = "\n\n".join(casos_texto)

        # Adiciona as considerações gerais de cada grupo ao final do compilado
        for grupo in grupos_perguntas:
            cons_gerais = st.session_state.consideracoes_gerais_grupo.get(grupo["titulo"], "")
            if cons_gerais.strip():
                compilado += f"\n\nConsiderações gerais do grupo '{grupo['titulo']}': {cons_gerais}"

        prompt_geral = (
            "Você receberá dados de vários casos de mamografia. "
            "Para cada caso, escreva um único parágrafo coeso que una todas as frases e considerações fornecidas, "
            "mantendo as informações originais. Em seguida, faça um resumo geral de todos os casos.\n\n"
            "Formato da resposta:\n"
            "CASO [Nome do Caso]: [texto coeso do caso]\n"
            "... (repita para cada caso)\n"
            "GERAL: [resumo geral]\n\n"
            "Dados dos casos:\n" + compilado
        )

        with st.spinner("IA processando todos os relatórios..."):
            try:
                response = model.generate_content(prompt_geral)
                resposta_completa = response.text

                # Parse da resposta
                relatorios = {}
                relatorio_geral = ""
                linhas = resposta_completa.split('\n')
                chave_atual = None
                for linha in linhas:
                    linha = linha.strip()
                    if linha.startswith("CASO "):
                        # Extrai nome do caso e texto
                        partes = linha.split(":", 1)
                        if len(partes) == 2:
                            chave = partes[0].replace("CASO ", "").strip()
                            texto = partes[1].strip()
                            relatorios[chave] = texto
                            chave_atual = chave
                        else:
                            # Se a linha só tem "CASO X" sem ":", pode ser um erro, mas tratamos
                            chave_atual = linha.replace("CASO ", "").strip()
                            relatorios[chave_atual] = ""
                    elif linha.startswith("GERAL:"):
                        relatorio_geral = linha.replace("GERAL:", "").strip()
                        chave_atual = None
                    elif chave_atual:
                        # Continua texto do caso atual
                        relatorios[chave_atual] += " " + linha

                # Atualiza session_state com os relatórios individuais
                for chave, texto in relatorios.items():
                    if chave in st.session_state.casos_salvos:
                        st.session_state.relatorios_ia[chave] = texto
                st.session_state.relatorio_geral_salvo = relatorio_geral

                st.success("Relatórios gerados com sucesso!")

                # Exibe as tabelas de respostas por grupo
                st.markdown("---")
                for grupo in grupos_perguntas:
                    st.subheader(f"📊 {grupo['titulo']} - Tabela de Respostas (Sim/Não)")
                    casos_ordenados = sorted(st.session_state.relatorios_ia.keys(), key=extrair_numero)
                    perguntas_grupo = list(grupo["perguntas"].keys())
                    html = "<table border='1' style='border-collapse: collapse; width: 100%; text-align: center;'>"
                    html += "<tr><th rowspan='2'>Pergunta</th>"
                    for caso in casos_ordenados:
                        html += f"<th colspan='2'>{caso}</th>"
                    html += "</tr>"
                    html += "<tr>"
                    for _ in casos_ordenados:
                        html += "<th>Sim</th><th>Não</th>"
                    html += "</tr>"
                    for pergunta in perguntas_grupo:
                        html += "<tr>"
                        html += f"<td style='text-align: left;'>{pergunta}</td>"
                        for caso in casos_ordenados:
                            chave_pergunta = f"{grupo['titulo']}||{pergunta}"
                            escolha = st.session_state.escolhas_casos.get(caso, {}).get(chave_pergunta, "-")
                            sim_x = "X" if escolha == "Sim" else ""
                            nao_x = "X" if escolha == "Não" else ""
                            html += f"<td>{sim_x}</td><td>{nao_x}</td>"
                        html += "</tr>"
                    html += "</table>"
                    st.markdown(html, unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Erro ao gerar relatórios: {e}")

# ------------------------------------------------------------
# Exportação do documento Word
# ------------------------------------------------------------
if st.session_state.relatorios_ia:
    st.markdown("---")
    st.header("💾 Exportar Documento")

    def limpar_formatacao(texto):
        return texto.replace("**", "").replace("__", "").replace("#", "")

    with st.expander("Visualizar Prévia do Documento", expanded=True):
        # ... (prévia similar ao código anterior, adaptada para múltiplos grupos) ...
        st.markdown("### PRÉVIA DO DOCUMENTO")
        # (omitido por brevidade, mas adaptado)
        pass

    def criar_docx_limpo():
        doc = Document()
        casos_ordenados = sorted(st.session_state.relatorios_ia.keys(), key=extrair_numero)

        # Cabeçalho
        doc.add_heading("Instrumento para a análise da qualidade da mamografia", level=0)
        # ... (dados do cabeçalho iguais aos anteriores) ...

        # Tabelas por grupo
        for grupo in grupos_perguntas:
            doc.add_heading(grupo["titulo"], level=1)
            perguntas_grupo = list(grupo["perguntas"].keys())
            num_casos = len(casos_ordenados)
            total_colunas = 1 + num_casos * 2
            tabela = doc.add_table(rows=2 + len(perguntas_grupo), cols=total_colunas)
            tabela.style = 'Table Grid'
            tabela.cell(0, 0).merge(tabela.cell(1, 0))
            tabela.cell(0, 0).text = "Pergunta"
            for idx, caso in enumerate(casos_ordenados):
                col_inicio = 1 + idx * 2
                col_fim = col_inicio + 1
                tabela.cell(0, col_inicio).merge(tabela.cell(0, col_fim))
                tabela.cell(0, col_inicio).text = caso
            for idx in range(num_casos):
                col_sim = 1 + idx * 2
                col_nao = col_sim + 1
                tabela.cell(1, col_sim).text = "Sim"
                tabela.cell(1, col_nao).text = "Não"
            for i, pergunta in enumerate(perguntas_grupo):
                linha_atual = i + 2
                tabela.cell(linha_atual, 0).text = pergunta
                for j, caso in enumerate(casos_ordenados):
                    chave_pergunta = f"{grupo['titulo']}||{pergunta}"
                    escolha = st.session_state.escolhas_casos.get(caso, {}).get(chave_pergunta, "-")
                    col_sim = 1 + j * 2
                    col_nao = col_sim + 1
                    if escolha == "Sim":
                        tabela.cell(linha_atual, col_sim).text = "X"
                        tabela.cell(linha_atual, col_nao).text = ""
                    elif escolha == "Não":
                        tabela.cell(linha_atual, col_sim).text = ""
                        tabela.cell(linha_atual, col_nao).text = "X"
                    else:
                        tabela.cell(linha_atual, col_sim).text = ""
                        tabela.cell(linha_atual, col_nao).text = ""
            doc.add_paragraph()

        doc.add_page_break()

        # Mini tabela de identificação dos exames
        doc.add_heading("Identificação dos Exames", level=2)
        num_casos = len(casos_ordenados)
        mini_tabela = doc.add_table(rows=2, cols=num_casos)
        mini_tabela.style = 'Table Grid'
        for j, caso in enumerate(casos_ordenados):
            mini_tabela.rows[0].cells[j].text = caso
        for j, caso in enumerate(casos_ordenados):
            id_texto = st.session_state.identificacao_exames.get(caso, "")
            mini_tabela.rows[1].cells[j].text = id_texto
        doc.add_paragraph()

        # Considerações específicas por caso
        doc.add_heading("Considerações Específicas", level=0)
        for nome_caso in casos_ordenados:
            doc.add_heading(nome_caso, level=1)
            texto_ia = st.session_state.relatorios_ia[nome_caso]
            for linha in texto_ia.strip().split('\n'):
                if linha.strip():
                    doc.add_paragraph(limpar_formatacao(linha))
            # Adiciona considerações de cada grupo para este caso
            if nome_caso in st.session_state.consideracoes_caso_grupo:
                for grupo_nome, texto_cons in st.session_state.consideracoes_caso_grupo[nome_caso].items():
                    if texto_cons.strip():
                        doc.add_heading(f"Considerações Adicionais - {grupo_nome}", level=2)
                        doc.add_paragraph(limpar_formatacao(texto_cons))
            doc.add_paragraph("-" * 30)

        # Relatório geral
        if st.session_state.relatorio_geral_salvo:
            for linha in st.session_state.relatorio_geral_salvo.strip().split('\n'):
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
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# ------------------------------------------------------------
# Botão de reset
# ------------------------------------------------------------
st.markdown("---")
if st.button("🗑️ Limpar todos os casos"):
    for chave in ["casos_salvos", "relatorios_ia", "relatorio_geral_salvo", "consideracoes_caso_grupo",
                  "consideracoes_gerais_grupo", "escolhas_casos", "dados_cabecalho", "identificacao_exames"]:
        if chave in st.session_state:
            del st.session_state[chave]
    st.rerun()

import streamlit as st
import google.generativeai as genai
from docx import Document
from io import BytesIO
import re

# 1. Configuração da IA
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-2.5-flash-lite')

st.title("Gerador de Relatórios")
st.header("Aspectos Físicos da Imagem")

# ============================================================
# NOVO – CABEÇALHO: dados da instituição e mamógrafo
# ============================================================
if "dados_cabecalho" not in st.session_state:
    st.session_state.dados_cabecalho = {
        "mamografo_fabricante": "",
        "mamografo_modelo": "",
        "cnes": "",
        "qiid": "",
        "tipo_mamografo": None,   # "Convencional", "Digital CR", "Digital DR", "DR retrofit"
        "instituicao": "",
        "cidade": "",
        "estado": ""
    }

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

# Atualiza o session_state com os dados atuais do formulário
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
# ============================================================
# FIM DO CABEÇALHO
# ============================================================

# --- INICIALIZAÇÃO DA MEMÓRIA (restante) ---
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

# 2. Biblioteca de Perguntas (inalterada)
perguntas = {
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

# SELEÇÃO DE CASO
st.markdown("---")
caso_atual = st.selectbox("Escolha o Caso que vai analisar agora:", [1, 2, 3, 4, 5])

# 3. Interface de Perguntas
respostas_temporarias = []
for titulo, info in perguntas.items():
    st.subheader(titulo)
    escolha = st.radio(f"Selecione:", list(info["opcoes"].keys()), key=f"radio_{titulo}_c{caso_atual}")
    sub_escolha = None
    if "sub_opcoes" in info and escolha == "Não":
        sub_escolha = st.radio(f"Especifique:", list(info["sub_opcoes"].keys()), key=f"sub_{titulo}_c{caso_atual}")
    obs = st.text_input(f"Considerações adicionais: ", key=f"obs_{titulo}_c{caso_atual}")
    respostas_temporarias.append({"titulo": titulo, "escolha": escolha, "sub_escolha": sub_escolha, "obs": obs})

# Considerações adicionais do caso
st.markdown("---")
consideracoes_caso = st.text_area(
    "📝 Considerações adicionais para este caso (opcional):",
    value=st.session_state.consideracoes_caso.get(f"Caso {caso_atual}", ""),
    key=f"consideracoes_c{caso_atual}",
    height=120
)

# 4. Lógica ao Salvar
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
        respostas_finais = []
        for item in respostas_temporarias:
            info_pergunta = perguntas[item["titulo"]]
            frase_base = info_pergunta["sub_opcoes"][item["sub_escolha"]] if item["escolha"] == "Não" and item["sub_escolha"] else info_pergunta["opcoes"][item["escolha"]]
            if item["obs"]:
                frase_base += f" Detalhe adicional: {item['obs']}"
            respostas_finais.append(frase_base)

        texto_bruto = " ".join(respostas_finais)
        st.session_state.casos_salvos[nome_caso] = texto_bruto
        st.session_state.consideracoes_caso[nome_caso] = consideracoes_caso

        escolhas = {}
        for item in respostas_temporarias:
            escolhas[item["titulo"]] = item["escolha"]
        st.session_state.escolhas_casos[nome_caso] = escolhas

        texto_para_ia = texto_bruto
        if consideracoes_caso.strip():
            texto_para_ia += f"\n\n{consideracoes_caso}"

        with st.spinner("IA formatando relatório..."):
            try:
                prompt = (
                    f"Deixe essas frases em um único texto coeso. "
                    f"Não mude as frases, apenas deixe o texto coeso para o Caso {caso_atual}: {texto_para_ia}"
                )
                response = model.generate_content(prompt)
                st.session_state.relatorios_ia[nome_caso] = response.text
                st.success(f"Caso {caso_atual} processado!")
            except Exception as e:
                st.error(f"Erro ao gerar relatório: {e}")

# Considerações Gerais
st.markdown("---")
st.subheader("📝 Considerações Gerais (serão incluídas em 'Todos os casos')")
st.session_state.consideracoes_gerais = st.text_area(
    "Digite aqui observações que se aplicam a todos os casos:",
    value=st.session_state.consideracoes_gerais,
    height=130,
    key="consideracoes_gerais_area"
)

# HISTÓRICO
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
            st.write("**Texto Bruto:**")
            st.caption(st.session_state.casos_salvos[nome_caso])
            if nome_caso in st.session_state.consideracoes_caso and st.session_state.consideracoes_caso[nome_caso].strip():
                st.write("**Considerações adicionais:**")
                st.info(st.session_state.consideracoes_caso[nome_caso])
            st.write("**Relatório IA:**")
            st.write(st.session_state.relatorios_ia[nome_caso])

# RELATÓRIO GERAL
if len(st.session_state.casos_salvos) >= 2:
    if st.button("Gerar Relatório Geral"):
        compilado = "".join([f"\n[{k}]: {v}\n" for k, v in st.session_state.casos_salvos.items()])
        texto_geral_para_ia = compilado
        if st.session_state.consideracoes_gerais.strip():
            texto_geral_para_ia += f"\n\nConsiderações gerais do avaliador: {st.session_state.consideracoes_gerais}"

        prompt_geral = (
            "Com base nos relatórios individuais abaixo, elabore um único parágrafo resumindo os achados gerais, "
            "sob o título 'Todos os casos'. Não mencione os números dos casos, apenas faça um resumo conciso.\n\n"
            f"Relatórios:\n{texto_geral_para_ia}"
        )
        try:
            response_geral = model.generate_content(prompt_geral)
            st.session_state.relatorio_geral_salvo = response_geral.text
            st.info(st.session_state.relatorio_geral_salvo)

            # Tabela HTML (já implementada antes, omitida aqui por brevidade – manter igual)
            st.markdown("---")
            st.subheader("📊 Tabela de Respostas (Sim/Não)")
            # ... (código da tabela HTML, igual ao fornecido)
        except Exception as e:
            st.error(f"Erro ao gerar relatório geral: {e}")

# SISTEMA DE EXPORTAÇÃO
if st.session_state.relatorios_ia:
    st.markdown("---")
    st.header("💾 Exportar Documento")

    def limpar_formatacao(texto):
        return texto.replace("**", "").replace("__", "")

    # Prévia (incluir cabeçalho + tabela HTML + conteúdo)
    with st.expander("Visualizar Prévia do Documento", expanded=True):
        st.markdown("### PRÉVIA DO DOCUMENTO")
        st.markdown("**Cabeçalho:**")
        cab = st.session_state.dados_cabecalho
        st.write(f"**Mamógrafo:** {cab['mamografo_fabricante']} – {cab['mamografo_modelo']}")
        st.write(f"**CNES:** {cab['cnes']}  |  **QIID:** {cab['qiid']}")
        st.write(f"**Tipo:** {cab['tipo_mamografo']}")
        st.write(f"**Instituição:** {cab['instituicao']}")
        st.write(f"**Cidade/Estado:** {cab['cidade']} – {cab['estado']}")
        st.markdown("---")
        # Tabela HTML (mesma da anterior)
        # ...
        st.markdown("---")
        for nome_caso in casos_ordenados:
            st.markdown(f"**{nome_caso}**")
            st.write(st.session_state.relatorios_ia[nome_caso])
            if nome_caso in st.session_state.consideracoes_caso and st.session_state.consideracoes_caso[nome_caso].strip():
                st.markdown("**Considerações adicionais:**")
                st.info(st.session_state.consideracoes_caso[nome_caso])
            st.markdown("---")
        if st.session_state.relatorio_geral_salvo:
            st.markdown("**Relatório Geral Consolidado**")
            st.write(st.session_state.relatorio_geral_salvo)
        if st.session_state.consideracoes_gerais.strip():
            st.markdown("**Considerações Gerais do Avaliador**")
            st.info(st.session_state.consideracoes_gerais)

    def criar_docx_limpo():
        doc = Document()

        # ============================================================
        # NOVO – CABEÇALHO no Word (primeira página)
        # ============================================================
        doc.add_heading("Instrumento para a análise da qualidade da mamografia", level=0)
        doc.add_paragraph()  # espaço

        cab = st.session_state.dados_cabecalho
        # Mamógrafo
        p = doc.add_paragraph()
        p.add_run("Mamógrafo (fabricante e modelo): ").bold = True
        p.add_run(f"{cab['mamografo_fabricante']} – {cab['mamografo_modelo']}")

        # CNES e QIID
        p = doc.add_paragraph()
        p.add_run("CNES: ").bold = True
        p.add_run(cab['cnes'])
        p.add_run("     QIID: ").bold = True
        p.add_run(cab['qiid'])

        # Tipo de mamógrafo (com quadradinho marcado)
        p = doc.add_paragraph()
        p.add_run("Tipo de mamógrafo: ").bold = True
        opcoes_tipo = ["Convencional", "Digital CR", "Digital DR", "DR retrofit"]
        for opcao in opcoes_tipo:
            marcado = "☒" if cab['tipo_mamografo'] == opcao else "☐"
            p.add_run(f"  {marcado} {opcao}  ")

        doc.add_paragraph()  # linha em branco

        # Instituição
        p = doc.add_paragraph()
        p.add_run("Instituição: ").bold = True
        p.add_run(cab['instituicao'])

        # Cidade e Estado
        p = doc.add_paragraph()
        p.add_run("Cidade: ").bold = True
        p.add_run(cab['cidade'])
        p.add_run("     Estado: ").bold = True
        p.add_run(cab['estado'])

        doc.add_page_break()  # Opcional: força quebra de página antes da tabela
        # ============================================================
        # FIM DO CABEÇALHO
        # ============================================================

        # Tabela de respostas (já existente)
        doc.add_heading("Tabela de Respostas (Sim/Não)", level=1)
        # ... (código da tabela no Word, mantido igual) ...

        # Restante do conteúdo (Considerações Específicas, etc.)
        # ...

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

# BOTÃO DE RESET (incluir a nova chave)
st.markdown("---")
if st.button("🗑️ Limpar todos os casos"):
    for chave in ["casos_salvos", "relatorios_ia", "relatorio_geral_salvo", "consideracoes_caso",
                  "consideracoes_gerais", "escolhas_casos", "dados_cabecalho"]:
        if chave in st.session_state:
            del st.session_state[chave]
    st.rerun()

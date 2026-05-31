import streamlit as st
import google.generativeai as genai
from docx import Document
from io import BytesIO
import re  # Para ordenação numérica dos casos
import pandas as pd  # Para gerenciar a visualização em tabela

# 1. Configuração da IA
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-2.5-flash-lite')

st.title("Gerador de Relatórios")

# Subtítulo fixo
st.header("Aspectos Físicos da Imagem")

# --- INICIALIZAÇÃO DA MEMÓRIA ---
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
if "tabela_respostas" not in st.session_state:
    st.session_state.tabela_respostas = {}

# 2. Biblioteca de Perguntas
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
                   "Não": "A áread e fundo da imagem não está adequadamente escura."},
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

# Campo para considerações adicionais do caso
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
        dicionario_respostas_brutas = {}
        
        for item in respostas_temporarias:
            info_pergunta = perguntas[item["titulo"]]
            dicionario_respostas_brutas[item["titulo"]] = item["escolha"]
            
            frase_base = info_pergunta["sub_opcoes"][item["sub_escolha"]] if item["escolha"] == "Não" and item["sub_escolha"] else info_pergunta["opcoes"][item["escolha"]]
            if item["obs"]:
                frase_base += f" Detalhe adicional: {item['obs']}"
            respostas_finais.append(frase_base)

        texto_bruto = " ".join(respostas_finais)
        st.session_state.casos_salvos[nome_caso] = texto_bruto
        st.session_state.consideracoes_caso[nome_caso] = consideracoes_caso
        st.session_state.tabela_respostas[nome_caso] = dicionario_respostas_brutas

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

# HISTÓRICO NA TELA
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

# Campo para Considerações Gerais
st.markdown("---")
st.subheader("📝 Considerações Gerais (serão incluídas em 'Todos os casos')")
st.session_state.consideracoes_generais = st.text_area(
    "Digite aqui observações que se aplicam a todos os casos:",
    value=st.session_state.consideracoes_generais,
    height=130,
    key="consideracoes_gerais_area"
)

# --- RELATÓRIO GERAL E QUADRO DE RESPOSTAS ---
if len(st.session_state.casos_salvos) >= 2:
    if st.button("Gerar Relatório Geral"):
        compilado = "".join([f"\n[{k}]: {v}\n" for k, v in st.session_state.casos_salvos.items()])
        texto_geral_para_ia = compilado
        if st.session_state.consideracoes_generais.strip():
            texto_geral_para_ia += f"\n\nConsiderações gerais do avaliador: {st.session_state.consideracoes_generais}"

        prompt_geral = (
            "Com base nos relatórios individuais abaixo, elabore um único parágrafo resumindo os achados gerais, "
            "sob o título 'Todos os casos'. Não mencione os números dos casos, apenas faça um resumo conciso.\n\n"
            f"Relatórios:\n{texto_geral_para_ia}"
        )
        try:
            response_geral = model.generate_content(prompt_geral)
            st.session_state.relatorio_geral_salvo = response_geral.text
            st.success("Relatório Geral gerado com sucesso!")
        except Exception as e:
            st.error(f"Erro ao gerar relatório geral: {e}")

# Exibe os resultados finais unificados (Tabela + Resumos) apenas se o relatório geral já tiver sido clicado/criado
if st.session_state.relatorio_geral_salvo:
    st.markdown("---")
    st.header("📊 Fechamento do Relatório Geral")
    
    # 1. Exibe a tabela na tela
    st.subheader("Quadro Comparativo de Respostas Principais")
    df_visualizacao = pd.DataFrame(st.session_state.tabela_respostas)
    colunas_ordenadas = sorted(df_visualizacao.columns, key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0)
    df_visualizacao = df_visualizacao[colunas_ordenadas]
    st.dataframe(df_visualizacao, use_container_width=True)
    
    # 2. Exibe o resumo gerado pela IA na tela
    st.subheader("Resumo Consolidado IA")
    st.info(st.session_state.relatorio_geral_salvo)

# SISTEMA DE EXPORTAÇÃO
if st.session_state.relatorios_ia:
    st.markdown("---")
    st.header("💾 Exportar Documento")

    def limpar_formatacao(texto):
        return texto.replace("**", "").replace("__", "")

    with st.expander("Visualizar Prévia Estrutural do Documento"):
        for nome_caso in casos_ordenados:
            st.markdown(f"**{nome_caso}**")
            st.write(st.session_state.relatorios_ia[nome_caso])
            st.markdown("---")

    def criar_docx_limpo():
        doc = Document()
        doc.add_heading("Considerações Específicas", 0)

        # --- AQUI É INJETADA A TABELA COMO A PRIMEIRA COISA DO ARQUIVO WORD ---
        if st.session_state.tabela_respostas:
            doc.add_heading("Quadro Comparativo de Análise", level=1)
            
            # Prepara os dados ordenados para montar as células
            df_doc = pd.DataFrame(st.session_state.tabela_respostas)
            colunas_doc = sorted(df_doc.columns, key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0)
            df_doc = df_doc[colunas_doc]
            
            # Cria a tabela no Word: número de linhas = itens + 1 (cabeçalho). Colunas = casos + 1 (aspectos)
            tabela_word = doc.add_table(rows=len(df_doc) + 1, cols=len(colunas_doc) + 1)
            tabela_word.style = 'Light Shading Accent 1'  # Estilo limpo padrão do Word
            
            # Preenche o cabeçalho
            hdr_cells = tabela_word.rows[0].cells
            hdr_cells[0].text = "Aspectos Físicos Avaliados"
            for idx, nome_col in enumerate(colunas_doc):
                hdr_cells[idx + 1].text = nome_col
                
            # Preenche os dados de cada linha
            for r_idx, aspecto in enumerate(df_doc.index):
                row_cells = tabela_word.rows[r_idx + 1].cells
                row_cells[0].text = aspecto
                for c_idx, nome_col in enumerate(colunas_doc):
                    row_cells[c_idx + 1].text = str(df_doc.at[aspecto, nome_col])
            
            doc.add_paragraph("\n" * 2) # Espaçamento depois da tabela

        # Segue a gravação dos textos individuais da IA
        for nome_caso in casos_ordenados:
            doc.add_heading(nome_caso, level=1)
            texto_ia = st.session_state.relatorios_ia[nome_caso]
            for linha in texto_ia.strip().split('\n'):
                if linha.strip():
                    doc.add_paragraph(limpar_formatacao(linha))
            if nome_caso in st.session_state.consideracoes_caso and st.session_state.consideracoes_caso[nome_caso].strip():
                doc.add_heading("Considerações Adicionais", level=2)
                doc.add_paragraph(limpar_formatacao(st.session_state.consideracoes_caso[nome_caso]))
            doc.add_paragraph("-" * 30)

        if st.session_state.relatorio_geral_salvo:
            doc.add_heading("Todos os Casos", level=1)
            for linha in st.session_state.relatorio_geral_salvo.strip().split('\n'):
                if linha.strip():
                    doc.add_paragraph(limpar_formatacao(linha))

        if st.session_state.consideracoes_generais.strip():
            doc.add_heading("Considerações Gerais do Avaliador", level=1)
            doc.add_paragraph(limpar_formatacao(st.session_state.consideracoes_generais))

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

# BOTÃO DE RESET
st.markdown("---")
if st.button("🗑️ Limpar todos os casos"):
    for chave in ["casos_salvos", "relatorios_ia", "relatorio_geral_salvo", "consideracoes_caso", "consideracoes_gerais", "tabela_respostas"]:
        if chave in st.session_state:
            del st.session_state[chave]
    st.rerun()

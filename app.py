import streamlit as st
import google.generativeai as genai
from docx import Document
from io import BytesIO
import re

# 1. Configuração da IA
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-2.5-flash-lite')

st.title("Gerador de Relatórios Multi-Casos")

# --- INICIALIZAÇÃO DA MEMÓRIA ---
if "casos_salvos" not in st.session_state:
    st.session_state.casos_salvos = {}
if "relatorios_ia" not in st.session_state:
    st.session_state.relatorios_ia = {}
if "relatorio_geral_salvo" not in st.session_state:
    st.session_state.relatorio_geral_salvo = None
# NOVO: armazena as considerações adicionais de cada caso separadamente
if "consideracoes_caso" not in st.session_state:
    st.session_state.consideracoes_caso = {}
# NOVO: considerações gerais (uma única string)
if "consideracoes_gerais" not in st.session_state:
    st.session_state.consideracoes_gerais = ""

# 2. Biblioteca de Perguntas (inalterada)
perguntas = {
    "Contraste adequado": {
        "opcoes": {"Sim": "O contraste está adequado.", "Não": "O contraste não está adequado."},
        "sub_opcoes": {
            "Contraste alto demais": "O contraste da imagem está alto demais.",
            "Contraste baixo demais": "O contraste está baixo demais."
        }
    },
    # ... (demais perguntas mantidas, sem alteração)
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
    obs = st.text_input(f"Detalhe (opcional):", key=f"obs_{titulo}_c{caso_atual}")
    respostas_temporarias.append({"titulo": titulo, "escolha": escolha, "sub_escolha": sub_escolha, "obs": obs})

# NOVO: campo para considerações adicionais do caso (aparece após todas as perguntas)
st.markdown("---")
consideracoes_caso = st.text_area(
    "📝 Considerações adicionais para este caso (opcional):",
    value=st.session_state.consideracoes_caso.get(f"Caso {caso_atual}", ""),
    key=f"consideracoes_c{caso_atual}",
    height=120
)

# 4. Lógica ao Salvar (modificada para incluir as considerações do caso)
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
        
        # NOVO: salva o texto bruto e também as considerações do caso
        st.session_state.casos_salvos[nome_caso] = texto_bruto
        st.session_state.consideracoes_caso[nome_caso] = consideracoes_caso  # armazena a string

        # Prepara o prompt incluindo as considerações (se houver)
        prompt_texto = texto_bruto
        if consideracoes_caso.strip():
            prompt_texto += f"\n\nConsiderações adicionais do avaliador: {consideracoes_caso}"

        with st.spinner("IA formatando relatório..."):
            try:
                prompt = f"Deixe essas frases em um único texto coeso. Não mude as frases, apenas deixe o texto coeso para o Caso {caso_atual}: {prompt_texto}"
                response = model.generate_content(prompt)
                st.session_state.relatorios_ia[nome_caso] = response.text
                st.success(f"Caso {caso_atual} processado!")
            except Exception as e:
                st.error(f"Erro ao gerar relatório: {e}")

# NOVO: campo para Considerações Gerais (aparece sempre abaixo do histórico)
st.markdown("---")
st.subheader("📝 Considerações Gerais (para o relatório consolidado)")
st.session_state.consideracoes_gerais = st.text_area(
    "Digite aqui observações que se aplicam a todos os casos:",
    value=st.session_state.consideracoes_gerais,
    height=130,
    key="consideracoes_gerais_area"
)

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
            # NOVO: exibe as considerações do caso se existirem
            if nome_caso in st.session_state.consideracoes_caso and st.session_state.consideracoes_caso[nome_caso].strip():
                st.write("**Considerações adicionais:**")
                st.info(st.session_state.consideracoes_caso[nome_caso])
            st.write("**Relatório IA:**")
            st.write(st.session_state.relatorios_ia[nome_caso])

# RELATÓRIO GERAL (agora incorpora as Considerações Gerais)
if len(st.session_state.casos_salvos) >= 2:
    if st.button("Gerar Relatório Geral"):
        compilado = "".join([f"\n[{k}]: {v}\n" for k, v in st.session_state.casos_salvos.items()])
        # NOVO: adiciona as considerações gerais ao prompt, se preenchidas
        prompt_geral = (
            "Com base nos relatórios individuais abaixo, elabore um único parágrafo resumindo os achados gerais, "
            "sob o título 'Todos os casos'. Não mencione os números dos casos, apenas faça um resumo conciso.\n\n"
            f"Relatórios:\n{compilado}"
        )
        if st.session_state.consideracoes_gerais.strip():
            prompt_geral += f"\n\nConsiderações gerais do avaliador: {st.session_state.consideracoes_gerais}"
        
        try:
            response_geral = model.generate_content(prompt_geral)
            st.session_state.relatorio_geral_salvo = response_geral.text
            st.info(st.session_state.relatorio_geral_salvo)
        except Exception as e:
            st.error(f"Erro ao gerar relatório geral: {e}")

# SISTEMA DE EXPORTAÇÃO
if st.session_state.relatorios_ia:
    st.markdown("---")
    st.header("💾 Exportar Documento")

    def limpar_formatacao(texto):
        return texto.replace("**", "").replace("__", "")

    with st.expander("Visualizar Prévia do Documento", expanded=True):
        st.markdown("### PRÉVIA DO DOCUMENTO (Como ficará no Word)")
        for nome_caso in casos_ordenados:
            st.markdown(f"**{nome_caso}**")
            st.write(st.session_state.relatorios_ia[nome_caso])
            # NOVO: mostra considerações do caso na prévia
            if nome_caso in st.session_state.consideracoes_caso and st.session_state.consideracoes_caso[nome_caso].strip():
                st.markdown("**Considerações adicionais:**")
                st.info(st.session_state.consideracoes_caso[nome_caso])
            st.markdown("---")
        if st.session_state.relatorio_geral_salvo:
            st.markdown("**Relatório Geral Consolidado**")
            st.write(st.session_state.relatorio_geral_salvo)
        # NOVO: mostra considerações gerais na prévia, se houver
        if st.session_state.consideracoes_gerais.strip():
            st.markdown("**Considerações Gerais do Avaliador**")
            st.info(st.session_state.consideracoes_gerais)

    def criar_docx_limpo():
        doc = Document()
        doc.add_heading("Relatório Técnico Consolidado", 0)

        for nome_caso in casos_ordenados:
            doc.add_heading(nome_caso, level=1)
            texto_ia = st.session_state.relatorios_ia[nome_caso]
            for linha in texto_ia.strip().split('\n'):
                if linha.strip():
                    doc.add_paragraph(limpar_formatacao(linha))
            # NOVO: adiciona as considerações do caso no documento
            if nome_caso in st.session_state.consideracoes_caso and st.session_state.consideracoes_caso[nome_caso].strip():
                doc.add_heading("Considerações Adicionais", level=2)
                doc.add_paragraph(limpar_formatacao(st.session_state.consideracoes_caso[nome_caso]))
            doc.add_paragraph("-" * 30)

        if st.session_state.relatorio_geral_salvo:
            doc.add_heading("Relatório Geral de Encerramento", level=1)
            for linha in st.session_state.relatorio_geral_salvo.strip().split('\n'):
                if linha.strip():
                    doc.add_paragraph(limpar_formatacao(linha))
        
        # NOVO: adiciona as considerações gerais ao final do documento
        if st.session_state.consideracoes_gerais.strip():
            doc.add_heading("Considerações Gerais do Avaliador", level=1)
            doc.add_paragraph(limpar_formatacao(st.session_state.consideracoes_gerais))

        output = BytesIO()
        doc.save(output)
        output.seek(0)
        return output

    st.download_button(
        label="📥 Baixar Documento Final (.docx)",
        data=criar_docx_limpo(),
        file_name="relatorio_final.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# BOTÃO DE RESET (agora também limpa as novas chaves)
st.markdown("---")
if st.button("🗑️ Limpar todos os casos"):
    for chave in ["casos_salvos", "relatorios_ia", "relatorio_geral_salvo", "consideracoes_caso", "consideracoes_gerais"]:
        if chave in st.session_state:
            del st.session_state[chave]
    st.rerun()

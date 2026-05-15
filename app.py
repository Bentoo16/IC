import streamlit as st
import google.generativeai as genai
from docx import Document
from io import BytesIO

# 1. Configuração da IA 
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-2.5-flash-lite')

st.title("📝 Gerador de Relatórios Multi-Casos")

# --- INICIALIZAÇÃO DA MEMÓRIA ---
if "casos_salvos" not in st.session_state:
    st.session_state.casos_salvos = {}
if "relatorios_ia" not in st.session_state:
    st.session_state.relatorios_ia = {}
if "relatorio_geral_salvo" not in st.session_state:
    st.session_state.relatorio_geral_salvo = None

# 2. Biblioteca de Perguntas
perguntas = {
    "Contraste adequado": {
        "opcoes": {"Sim": "O contraste está adequado.", "Não": "O contraste não está adequado."},
        "sub_opcoes": {
            "Contraste alto demias": "O contraste da imagem está alto demais.",
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
            "Problema A": "Frase gerada para o problema .",
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
caso_atual = st.selectbox(" Escolha o Caso que vai analisar agora:", [1, 2, 3, 4, 5])

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

# 4. Lógica ao Salvar 
if st.button(f" Analisar e Salvar Caso {caso_atual}"):
    respostas_finais = []
    for item in respostas_temporarias:
        info_pergunta = perguntas[item["titulo"]]
        frase_base = info_pergunta["sub_opcoes"][item["sub_escolha"]] if item["escolha"] == "Não" and item["sub_escolha"] else info_pergunta["opcoes"][item["escolha"]]
        if item["obs"]: frase_base += f" Detalhe adicional: {item['obs']}"
        respostas_finais.append(frase_base)

    texto_bruto = " ".join(respostas_finais)
    st.session_state.casos_salvos[f"Caso {caso_atual}"] = texto_bruto
    
    # EXIBIÇÃO DO TEXTO BRUTO NA TELA 
    st.markdown(f"### 📋 Texto Bruto do Caso {caso_atual}:")
    st.warning(texto_bruto)

    with st.spinner("IA formatando relatório..."):
        prompt = f"Deixe essas frases em um único texto coeso. Não mude as frases, apenas deixe o texto coeso para o Caso {caso_atual}: {texto_bruto}"
        response = model.generate_content(prompt)
        st.session_state.relatorios_ia[f"Caso {caso_atual}"] = response.text
        st.success(f"Caso {caso_atual} processado!")

# HISTÓRICO NA TELA
if st.session_state.relatorios_ia:
    st.markdown("---")
    st.header(" Histórico da Sessão")
    abas = st.tabs(list(st.session_state.relatorios_ia.keys()))
    for i, nome_caso in enumerate(st.session_state.relatorios_ia.keys()):
        with abas[i]:
            # Texto bruto continua visível aqui nas abas
            st.write("**Texto Bruto:**")
            st.caption(st.session_state.casos_salvos[nome_caso])
            st.write("**Relatório IA:**")
            st.write(st.session_state.relatorios_ia[nome_caso])

# RELATÓRIO GERAL 
if len(st.session_state.casos_salvos) >= 2:
    if st.button(" Gerar Relatório Geral"):
        compilado = "".join([f"\n[{k}]: {v}\n" for k, v in st.session_state.casos_salvos.items()])
        response_geral = model.generate_content(f"Faça um relatório geral dos casos: {compilado}")
        st.session_state.relatorio_geral_salvo = response_geral.text
        st.info(st.session_state.relatorio_geral_salvo)

# SISTEMA DE EXPORTAÇÃO 
if st.session_state.relatorios_ia:
    st.markdown("---")
    st.header("💾 Exportar Documento")

    def limpar_formatacao(texto):
        # Remove os asteriscos (**) que a IA usa para negrito
        return texto.replace("**", "").replace("__", "")
    # --- BLOCO DA PRÉVIA (Apenas visual no Streamlit) ---
    with st.expander(" Visualizar Prévia do Documento", expanded=True):
        st.markdown("### PRÉVIA DO DOCUMENTO (Como ficará no Word)")
        for nome_caso in sorted(st.session_state.relatorios_ia.keys()):
            st.markdown(f"**{nome_caso}**")
            # Na prévia mostramos o texto da IA (com negrito do markdown funcionando no st.write)
            st.write(st.session_state.relatorios_ia[nome_caso])
            st.markdown("---")
        if st.session_state.relatorio_geral_salvo:
            st.markdown("**Relatório Geral Consolidado**")
            st.write(st.session_state.relatorio_geral_salvo)

    def criar_docx_limpo():
        doc = Document()
        doc.add_heading("Relatório Técnico Consolidado", 0)

        for nome_caso in sorted(st.session_state.relatorios_ia.keys()):
            doc.add_heading(nome_caso, level=1)
            
            # ADICIONA APENAS O TEXTO DA IA 
            texto_ia = st.session_state.relatorios_ia[nome_caso]
            doc.add_paragraph(limpar_formatacao(texto_ia))
            
            doc.add_paragraph("-" * 30)

        if st.session_state.relatorio_geral_salvo:
            doc.add_heading("Relatório Geral de Encerramento", level=1)
            doc.add_paragraph(limpar_formatacao(st.session_state.relatorio_geral_salvo))

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

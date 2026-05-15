import streamlit as st
import google.generativeai as genai
from docx import Document  # Biblioteca para manipular arquivos do Word
from io import BytesIO  # Permite criar o arquivo na memória antes do download

# 1. Configuração da IA
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-flash-latest')

st.title("📝 Gerador de Relatórios Multi-Casos (Até 5 Casos)")

# --- INICIALIZAÇÃO DA MEMÓRIA (Session State) ---
if "casos_salvos" not in st.session_state:
    st.session_state.casos_salvos = {}  # Guarda o texto bruto de cada caso
if "relatorios_ia" not in st.session_state:
    st.session_state.relatorios_ia = {}  # Guarda a resposta da IA de cada caso
if "relatorio_geral_salvo" not in st.session_state:
    st.session_state.relatorio_geral_salvo = None  # Guarda o relatório geral gerado pela IA

# 2. Biblioteca de Perguntas
perguntas = {
    "Contraste adequado": {
        "opcoes": {"Sim": "O contraste está adequado.", "Não": "O contraste não está adequado."},
        "sub_opcoes": {
            "Problema genérico A": "Frase gerada para o problema A do contraste.",
            "Problema genérico B": "Frase gerada para o problema B do contraste."
        }
    },
    "Definição de estruturas": {
        "opcoes": {"Sim": "Está bem definido.", "Não": "Não está bem definido."},
        "sub_opcoes": {
            "Problema genérico A": "Frase gerada para o problema A da definição.",
            "Problema genérico B": "Frase gerada para o problema B da definição."
        }
    },
    "Saturação correta nas áreas claras": {
        "opcoes": {"Sim": "Está bem saturada nas áreas claras.", "Não": "Não está bem saturada nas áreas claras."},
        "sub_opcoes": {
            "Problema genérico A": "Frase gerada para o problema A das áreas claras.",
            "Problema genérico B": "Frase gerada para o problema B das áreas claras."
        }
    },
    "Saturação correta nas áreas escuras": {
        "opcoes": {"Sim": "Está bem saturada nas áreas escuras.", "Não": "Não está bem saturada nas áreas escuras."},
        "sub_opcoes": {
            "Tem muita coisa": "A imagem apresenta supersaturação (excesso) nas áreas escuras.",
            "Tem pouca coisa": "A imagem apresenta sub-saturação (falta) nas áreas escuras."
        }
    },
    "Imagem sem ruído": {
        "opcoes": {"Sim": "Está sem ruído.", "Não": "Está com ruído."},
        "sub_opcoes": {
            "Problema genérico A": "Frase gerada para o problema A de ruído.",
            "Problema genérico B": "Frase gerada para o problema B de ruído."
        }
    },
    "A área de fundo está adequadamente escura (enegrecimento película)": {
        "opcoes": {"Sim": "A área está adequadamente escura.", "Não": "A área não está adequadamente escura."},
        "sub_opcoes": {
            "Problema genérico A": "Frase gerada para o problema A do fundo.",
            "Problema genérico B": "Frase gerada para o problema B do fundo."
        }
    },
    "Imagem sem artefatos (se houver, descrever)": {
        "opcoes": {"Sim": "Não possui artefatos.", "Não": "Possui artefatos."},
        "sub_opcoes": {
            "Problema genérico A": "Frase gerada para o problema A de artefatos.",
            "Problema genérico B": "Frase gerada para o problema B de artefatos."
        }
    }
}

# --- SELEÇÃO DE QUAL CASO ESTÁ SENDO ANALISADO ---
st.markdown("---")
caso_atual = st.selectbox(" Escolha o Caso que vai analisar agora:", [1, 2, 3, 4, 5])
st.info(f"Modo de Edição: preenchendo dados para o **Caso {caso_atual}**")

# 3. Interface de Perguntas
respostas_temporarias = []

for titulo, info in perguntas.items():
    st.subheader(titulo)
    
    escolha = st.radio(f"Selecione:", list(info["opcoes"].keys()), key=f"radio_{titulo}_c{caso_atual}")

    sub_escolha = None
    if "sub_opcoes" in info and escolha == "Não":
        sub_escolha = st.radio(
            f"Especifique o problema para {titulo}:", 
            list(info["sub_opcoes"].keys()), 
            key=f"sub_{titulo}_c{caso_atual}"
        )

    obs = st.text_input(f"Descrever detalhe (opcional):", key=f"obs_{titulo}_c{caso_atual}")

    respostas_temporarias.append({
        "titulo": titulo, 
        "escolha": escolha, 
        "sub_escolha": sub_escolha, 
        "obs": obs
    })

st.markdown("---")
salvar_caso = st.button(f" Analisar e Salvar Caso {caso_atual}")

# 4. Lógica ao Salvar o Caso Atual
if salvar_caso:
    respostas_finais = []
    for item in respostas_temporarias:
        info_pergunta = perguntas[item["titulo"]]
        if item["escolha"] == "Não" and item["sub_escolha"]:
            frase_base = info_pergunta["sub_opcoes"][item["sub_escolha"]]
        else:
            frase_base = info_pergunta["opcoes"][item["escolha"]]

        if item["obs"]:
            frase_base += f" Detalhe adicional: {item['obs']}"
        respostas_finais.append(frase_base)

    texto_bruto_caso = " ".join(respostas_finais)
    st.session_state.casos_salvos[f"Caso {caso_atual}"] = texto_bruto_caso

    st.markdown("---")
    st.markdown(f"### 📋 Frases Juntas do Caso {caso_atual} (Texto Bruto):")
    st.warning(texto_bruto_caso)

    with st.spinner(f"Processando Relatório Individual do Caso {caso_atual}..."):
        prompt_individual = f"Deixe essas frases em um text coeso, não adicione nada além do que está nas frases, para o Caso {caso_atual}: {texto_bruto_caso}"
        try:
            response = model.generate_content(prompt_individual)
            st.session_state.relatorios_ia[f"Caso {caso_atual}"] = response.text
            
            st.markdown(f"### Relatório Individual do Caso {caso_atual} (Com IA):")
            st.success(response.text)
        except Exception as e:
            st.error(f"Erro na IA: {e}")

# --- PAINEL DE HISTÓRICO DE RESULTADOS ---
if st.session_state.relatorios_ia:
    st.markdown("---")
    st.header(" Histórico de Casos Salvos nesta Sessão")
    
    abas = st.tabs(list(st.session_state.relatorios_ia.keys()))
    for i, nome_caso in enumerate(st.session_state.relatorios_ia.keys()):
        with abas[i]:
            st.text_area("Texto Bruto Salvo:", st.session_state.casos_salvos[nome_caso], height=70, disabled=True, key=f"area_{nome_caso}")
            st.write(st.session_state.relatorios_ia[nome_caso])

# --- GERAÇÃO DO RELATÓRIO GERAL (Consolidado) ---
if len(st.session_state.casos_salvos) >= 2:
    st.markdown("---")
    st.header(" Fechamento do Relatório Geral")
    st.write(f"Você já tem {len(st.session_state.casos_salvos)} casos prontos para consolidar.")
    
    gerar_geral = st.button(" Gerar Relatório Geral Consolidado")
    
    if gerar_geral:
        compilado_todos_casos = ""
        for nome_caso, texto_caso in st.session_state.casos_salvos.items():
            compilado_todos_casos += f"\n[{nome_caso}]: {texto_caso}\n"
            
        st.markdown(f"### Compilado de todos os casos (Texto Bruto):")
        st.warning(compilado_todos_casos)
        
        with st.spinner("O Gemini está cruzando os dados e redigindo o relatório geral..."):
            prompt_geral = f"""
            Dado todos os casos, agora faça um relatório geral ressaltando os pontos importantes.
            
            Dados dos casos:
            {compilado_todos_casos}
            """
            try:
                response_geral = model.generate_content(prompt_geral)
                st.session_state.relatorio_geral_salvo = response_geral.text
                st.markdown("---")
                st.markdown("###  RELATÓRIO GERAL CONSOLIDADO:")
                st.info(st.session_state.relatorio_geral_salvo)
            except Exception as e:
                st.error(f"Erro ao gerar relatório geral: {e}")

# --- SEÇÃO EXCLUSIVA PARA EXPORTAÇÃO PARA WORD (NOVO) ---
if st.session_state.casos_salvos:
    st.markdown("---")
    st.header("💾 Exportar Resultados para Word")
    st.write("Clique no botão abaixo para gerar o arquivo `.docx` completo com todos os dados atuais.")
    
    # Função que monta o documento Word na memória
    def criar_arquivo_word():
        doc = Document()
        doc.add_heading("Relatório Técnico de Casos Múltiplos", level=0)
        
        # Percorre cada caso adicionando seus detalhes ao Word
        for nome_caso in sorted(st.session_state.casos_salvos.keys()):
            doc.add_heading(nome_caso, level=1)
            
            doc.add_heading("Texto Bruto Selecionado:", level=2)
            doc.add_paragraph(st.session_state.casos_salvos[nome_caso])
            
            if nome_caso in st.session_state.relatorios_ia:
                doc.add_heading("Relatório Individual (IA):", level=2)
                doc.add_paragraph(st.session_state.relatorios_ia[nome_caso])
                
            doc.add_paragraph("-" * 30) # Separador de linha simples no Word
            
        # Adiciona o Relatório Geral ao fim do documento, se ele existir
        if st.session_state.relatorio_geral_salvo:
            doc.add_heading("Relatório Geral Consolidado", level=1)
            doc.add_paragraph(st.session_state.relatorio_geral_salvo)
            
        # Salva o documento no buffer de memória BytesIO
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    # Executa a criação do Word e exibe o botão nativo de download do Streamlit
    arquivo_docx = criar_arquivo_word()
    
    st.download_button(
        label="📥 Baixar Arquivo Word (.docx)",
        data=arquivo_docx,
        file_name="relatorio_completo_casos.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

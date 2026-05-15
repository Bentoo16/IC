import streamlit as st
import google.generativeai as genai

# 1. Configuração da IA
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-flash-latest')

st.title("📝 Gerador de Relatórios Multi-Casos (Até 5 Casos)")

# --- INICIALIZAÇÃO DA MEMÓRIA (Session State) ---
if "casos_salvos" not in st.session_state:
    st.session_state.casos_salvos = {}  # Guarda o texto bruto de cada caso
if "relatorios_ia" not in st.session_state:
    st.session_state.relatorios_ia = {}  # Guarda a resposta da IA de cada caso

# 2. Biblioteca de Perguntas (TODAS com sub-opções genéricas condicionadas)
perguntas = {
    "Contraste adequado": {
        "opcoes": {"Sim": "O contraste está adequado.", "Não": "O contraste não está adequado."},
        "sub_opcoes": {
            "Contraste alto demias": "O contraste da imagem está alto demais.",
            "Contraste baixo demais": "O contraste está baixo demais."
        }
    },
    "Definição de estruturas": {
        "opcoes": {"Sim": "Está bem definido.", "Não": "Não está bem definido."},
        "sub_opcoes": {
            "Problema A": "Frase gerada para o problema A.",
            "Problema B": "Frase gerada para o problema B."
        }
    },
    "Saturação correta nas áreas claras": {
        "opcoes": {"Sim": "Está bem saturada nas áreas claras.", "Não": "Não está bem saturada nas áreas claras."},
        "sub_opcoes": {
            "Problema A": "Frase gerada para o problema A.",
            "Problema B": "Frase gerada para o problema B."
        }
    },
    "Saturação correta nas áreas escuras": {
        "opcoes": {"Sim": "Está bem saturada nas áreas escuras.", "Não": "Não está bem saturada nas áreas escuras."},
        "sub_opcoes": {
            "Problema A": "Frase gerada para o problema .",
            "Problema B": "Frase gerada para o problema B."
        }
    },
    "Imagem sem ruído": {
        "opcoes": {"Sim": "Está sem ruído.", "Não": "Está com ruído."},
        "sub_opcoes": {
            "Problema A": "Frase gerada para o problema A.",
            "Problema B": "Frase gerada para o problema B."
        }
    },
    "A área de fundo está adequadamente escura (enegrecimento película)": {
        "opcoes": {"Sim": "A área está adequadamente escura.", "Não": "A área não está adequadamente escura."},
        "sub_opcoes": {
            "Problema A": "Frase gerada para o problema A.",
            "Problema genérico B": "Frase gerada para o problema B."
        }
    },
    "Imagem sem artefatos (se houver, descrever)": {
        "opcoes": {"Sim": "Não possui artefatos.", "Não": "Possui artefatos."},
        "sub_opcoes": {
            "Problema A": "Frase gerada para o problema A.",
            "Problema B": "Frase gerada para o problema B."
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
                st.markdown("---")
                st.markdown("###  RELATÓRIO GERAL CONSOLIDADO:")
                st.info(response_geral.text)
            except Exception as e:
                st.error(f"Erro ao gerar relatório geral: {e}")

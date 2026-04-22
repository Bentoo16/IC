import streamlit as st
import google.generativeai as genai

# 1. Configuração
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-flash-latest')

st.title("📝 Gerador de Relatórios")

# 2. Biblioteca de Perguntas (CORRIGIDA)
perguntas = {
    "Contraste adequado": {
        "opcoes": {
            "Sim": "O contraste está adequado.",
            "Não": "O contraste não está adequado"
        }
    },
    "Definição de estruturas": {
        "opcoes": {
            "Sim": "Está bem definido.",
            "Não": "Não está bem definido."
        }
    },
    "Saturação correta nas áreas claras": {
        "opcoes": {
            "Sim": "Está bem saturada nas áreas claras.",
            "Não": "Não está bem saturada nas áreas claras."
        }
    },
    "Saturação correta nas áreas escuras": {
        "opcoes": {
            "Sim": "Está bem saturada nas áreas escuras.",
            "Não": "Não está bem saturada nas áreas escuras"
        }
    },
    "Imagem sem ruído": {
        "opcoes": {
            "Sim": "Está sem ruído.",
            "Não": "Está com ruído"
        }
    },
    "A área de fundo está adequadamente escura (enegrecimento película)": {
        "opcoes": {
            "Sim": "A área está adequadamente escura",
            "Não": "A área não está adequadamente escura"
        }
    },
    "Imagem sem artefatos (se houver, descrever)": {
        "opcoes": {
            "Sim": "Não possui artefatos",
            "Não": "Possui artefatos"
        }
    }
}

# 3. Interface
# Criamos a lista dentro do formulário para capturar os dados apenas no clique do botão
with st.form("relatorio_form"):
    respostas_temporarias = []

    for titulo, info in perguntas.items():
        st.subheader(titulo)

        # Opção fixa (Sim/Não)
        escolha = st.radio(f"Selecione:", list(info["opcoes"].keys()), key=f"radio_{titulo}")

        # Campo livre (Sempre visível)
        obs = st.text_input(f"Descrever detalhe (opcional):", key=f"obs_{titulo}")

        # Guardamos a escolha e a obs para processar depois do submit
        respostas_temporarias.append({"titulo": titulo, "escolha": escolha, "obs": obs})

    submit = st.form_submit_button("Gerar Relatório")

# 4. Processamento da IA (Só acontece após o submit)
if submit:
    respostas_finais = []

    for item in respostas_temporarias:
        # Pega a frase pronta baseada na escolha (Sim/Não)
        frase_base = perguntas[item["titulo"]]["opcoes"][item["escolha"]]

        # Adiciona a observação se ela existir
        if item["obs"]:
            frase_base += f" Detalhe adicional: {item['obs']}"

        respostas_finais.append(frase_base)

    texto_bruto = " ".join(respostas_finais)

    with st.spinner("Formatando relatório..."):
        prompt = f"Transforme estas frases em um relatório técnico profissional e coeso: {texto_bruto}"

        try:
            response = model.generate_content(prompt)
            st.markdown("---")
            st.markdown("### Resultado Final:")
            st.success(response.text)
        except Exception as e:
            st.error(f"Erro ao conectar com a IA: {e}")

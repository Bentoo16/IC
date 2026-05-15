import streamlit as st
import google.generativeai as genai

# 1. Configuração
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-flash-latest')

st.title("📝 Gerador de Relatórios")

# 2. Biblioteca de Perguntas
perguntas = {
    "Contraste adequado": {
        "opcoes": {
            "Sim": "O contraste está adequado.",
            "Não": "O contraste não está adequado."
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
        },
        "sub_opcoes": {
            "Tem muita coisa": "A imagem apresenta supersaturação (excesso) nas áreas claras.",
            "Tem pouca coisa": "A imagem apresenta sub-saturação (falta) nas áreas claras."
        }
    },
    "Saturação correta nas áreas escuras": {
        "opcoes": {
            "Sim": "Está bem saturada nas áreas escuras.",
            "Não": "Não está bem saturada nas áreas escuras."
        },
        "sub_opcoes": {
            "Tem muita coisa": "A imagem apresenta supersaturação (excesso) nas áreas escuras.",
            "Tem pouca coisa": "A imagem apresenta sub-saturação (falta) nas áreas escuras."
        }
    },
    "Imagem sem ruído": {
        "opcoes": {
            "Sim": "Está sem ruído.",
            "Não": "Está com ruído."
        }
    },
    "A área de fundo está adequadamente escura (enegrecimento película)": {
        "opcoes": {
            "Sim": "A área está adequadamente escura.",
            "Não": "A área não está adequadamente escura."
        }
    },
    "Imagem sem artefatos (se houver, descrever)": {
        "opcoes": {
            "Sim": "Não possui artefatos.",
            "Não": "Possui artefatos."
        }
    }
}

# 3. Interface (SEM st.form para permitir atualização em tempo real)
respostas_temporarias = []

for titulo, info in perguntas.items():
    st.subheader(titulo)

    # Opção fixa principal (Sim/Não)
    escolha = st.radio(f"Selecione:", list(info["opcoes"].keys()), key=f"radio_{titulo}")

    # LÓGICA CONDICIONAL INSTANTÂNEA
    sub_escolha = None
    if "sub_opcoes" in info and escolha == "Não":
        sub_escolha = st.radio(
            f"Especifique o problema para {titulo}:", 
            list(info["sub_opcoes"].keys()), 
            key=f"sub_{titulo}"
        )

    # Campo livre
    obs = st.text_input(f"Descrever detalhe (opcional):", key=f"obs_{titulo}")

    respostas_temporarias.append({
        "titulo": titulo, 
        "escolha": escolha, 
        "sub_escolha": sub_escolha, 
        "obs": obs
    })

st.markdown("---")
# O botão agora é um st.button comum
submit = st.button("Gerar Relatório")

# 4. Processamento da IA
if submit:
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

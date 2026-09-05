# 🦉 Prof. Aula

Bem-vindo ao **Prof. Aula**! Um assistente de mesa projetado especialmente para professores, focado em agilizar e otimizar a criação de planejamentos de aula e cadernos de atividades com o poder da Inteligência Artificial e alinhamento total à BNCC (Base Nacional Comum Curricular)[cite: 3].

## 📌 O que é o Prof. Aula?
O Prof. Aula é um aplicativo multiplataforma (funciona em Windows, macOS e Linux) que ajuda educadores a estruturarem suas aulas de forma rápida[cite: 3]. Em vez de gastar horas formatando documentos e buscando códigos da BNCC, o professor preenche os dados básicos da aula (ou dita por voz) e o aplicativo gera um documento no Word (.docx) pronto para impressão em layout otimizado[cite: 3].

---

## 🔑 Como Configurar as Chaves de API (IA)

O aplicativo utiliza um motor Multi-IA com fallback automático[cite: 3]. Para utilizá-lo, você precisa cadastrar ao menos uma chave de acesso (API Key) gratuita[cite: 3].

### 🟢 1. Google Gemini (Padrão & Recomendado)
O **Google Gemini é a IA principal e oficial do aplicativo**[cite: 3]. Ele é obrigatório como escolha padrão por entregar a melhor qualidade pedagógica, linguagem adequada para a realidade das turmas, excelente capacidade de interpretação de áudio e formação estrita de documentos[cite: 3].

1. Acesse o site oficial: [Google AI Studio](https://aistudio.google.com/app/apikey)[cite: 3]
2. Faça login com sua conta Google[cite: 3].
3. Clique no botão **"Create API Key"** (Criar chave de API)[cite: 3].
4. Copie o código gerado (iniciado por `AIzaSy...`)[cite: 3].
5. Abra o Prof. Aula, clique em **⚙️ Configurações Multi-IA**, cole no campo do Gemini e clique em **Salvar**[cite: 3].

---

### 🟡 2. Provedores de Backup e Segurança (Opcionais)
Caso a cota do Gemini oscile ou fique temporariamente congestionada, o aplicativo aciona automaticamente os provedores secundários para que você nunca fique sem gerar suas atividades[cite: 3].

* **Groq (Backup de Alta Velocidade):**[cite: 3]
  1. Acesse o [Console da Groq](https://console.groq.com/keys)[cite: 3].
  2. Faça login e clique em **"Create API Key"**[cite: 3].
  3. Copie a chave gerada (iniciada por `gsk_...`) e insira no campo dedicado em Configurações[cite: 3].

* **OpenRouter (Backup Multi-Modelo):**[cite: 3]
  1. Acesse o [OpenRouter Keys](https://openrouter.ai/workspaces/default/keys)[cite: 3].
  2. Crie uma chave e cole-a no campo OpenRouter do aplicativo[cite: 3].

---

## 🚀 Principais Funcionalidades
- **Geração Inteligente:** Criação de planos de aula e cadernos de exercícios via IA[cite: 3].
- **Ditação por Voz Integrada:** Digite instruções apenas falando no microfone[cite: 3].
- **Layout de Impressão Otimizado:** Controle de margens (estreitas), 2 colunas, linhas pontilhadas dinâmicas e seleção de tamanho de fonte para economizar papel[cite: 3].
- **Integração com a BNCC:** Filtro automático de habilidades alinhado à série selecionada[cite: 3].
- **Exportação Direta:** Gera arquivos `.docx` configurados diretamente no Microsoft Word[cite: 3].

---

## 📥 Como baixar e instalar

Acesse a nossa área de lançamentos oficiais para obter a versão mais recente compilada de forma estável:

👉 **[CLIQUE AQUI PARA BAIXAR A VERSÃO MAIS RECENTE](https://github.com/ramonling/profaula/releases/latest)** 👈

Escolha o arquivo correto para o seu sistema:

### 🪟 Para Windows
1. Baixe o arquivo **`ProfAula-Windows.exe`**.
2. Dê um duplo clique no arquivo baixado[cite: 3]. *(Se o Windows SmartScreen exibir um aviso de segurança por ser um app novo, clique em "Mais informações" e depois em "Executar assim mesmo")[cite: 3].*

### 🍎 Para macOS
1. Baixe o arquivo **`ProfAula-macOS.zip`**[cite: 3].
2. Extraia o arquivo ZIP[cite: 3].
3. Dê um duplo clique no aplicativo **ProfAula** para abri-lo. *(Se o Mac bloquear a abertura, vá em Ajustes do Sistema > Privacidade e Segurança e clique em "Abrir Mesmo Assim")[cite: 3].*

### 🐧 Para Linux
1. Baixe o arquivo **`ProfAula-Linux`**.
2. Clique com o botão direito no arquivo, vá em **Propriedades > Permissões** e marque a caixa para permitir a execução como um programa[cite: 3].
3. Dê um duplo clique para abrir[cite: 3].

---

## 🛠️ Para Desenvolvedores e Feedback
Este projeto é de código aberto e foi desenvolvido por Ramon utilizando:
- **Linguagem:** Python 3.11+[cite: 3]
- **Interface:** Tkinter (Customized UI)[cite: 3]
- **Bibliotecas:** `google-genai`, `python-docx`, `pypdf`, `SpeechRecognition`, `pyaudio`[cite: 3]
- **Empacotamento & CI/CD:** PyInstaller + GitHub Actions

💬 Encontrou um bug ou tem uma sugestão? Envie um e-mail para **profaulaai@gmail.com** ou acompanhe no Instagram **@ramonchvr**.

---
*Desenvolvido com 💻 e ☕ para facilitar a vida dos educadores.*[cite: 3]

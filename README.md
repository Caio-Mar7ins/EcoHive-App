# 🌿 EcoHive

> Aplicativo desktop de gestão de reciclagem para condomínios, desenvolvido em Python com interface gráfica moderna.

---

## 📋 Sumário

- [Sobre o Projeto](#sobre-o-projeto)
- [Funcionalidades](#funcionalidades)
- [Demonstração](#demonstração)
- [Tecnologias Utilizadas](#tecnologias-utilizadas)
- [Pré-requisitos](#pré-requisitos)
- [Instalação e Execução](#instalação-e-execução)
- [Estrutura do Projeto](#estrutura-do-projeto)
- [Arquitetura](#arquitetura)
- [Como Usar](#como-usar)
- [Armazenamento de Dados](#armazenamento-de-dados)
- [Contribuição](#contribuição)
- [Autor](#autor)
- [Licença](#licença)

---

## 📌 Sobre o Projeto

O **EcoHive** é um sistema desktop desenvolvido como trabalho acadêmico para a disciplina de Programação, com o objetivo de incentivar a reciclagem em condomínios residenciais por meio de um sistema de pontuação gamificado.

Moradores registram seus descartes de resíduos recicláveis e acumulam pontos que podem ser trocados por recompensas — como desconto na taxa condominial, vaga prioritária e vouchers. O sistema também exibe um painel coletivo com o total reciclado por bloco, promovendo engajamento da comunidade.

**Contexto:** Projeto acadêmico — UNIP, Ciência da Computação.

---

## ✅ Funcionalidades

- **Cadastro e Login de moradores** com validação de usuário/e-mail duplicado
- **Registro de descartes** por tipo de resíduo (Plástico, Papel, Vidro, Lata, Metal, Eletrônico) e peso em kg
- **Sistema de pontuação** automático por tipo e peso do resíduo descartado
- **Painel de bloco** com somatório total reciclado por categoria, filtrado pelo bloco do morador logado
- **Resgate de recompensas** com débito automático de pontos
- **Persistência de dados** em arquivo Excel local (`.xlsx`) com duas abas: `usuarios` e `descarte`
- **Interface responsiva** com imagens de fundo escaláveis e placeholders customizados nos campos
- **Alternância de telas** entre Login, Cadastro, Home (painel) e Resgate

---

## 🛠️ Tecnologias Utilizadas

| Tecnologia | Finalidade |
|---|---|
| Python 3.10+ | Linguagem principal |
| [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) | Interface gráfica moderna (tema escuro/claro) |
| [Pandas](https://pandas.pydata.org/) | Leitura, manipulação e gravação de dados em Excel |
| [Pillow (PIL)](https://python-pillow.org/) | Processamento e redimensionamento de imagens |
| [openpyxl](https://openpyxl.readthedocs.io/) | Engine de escrita para arquivos `.xlsx` |
| [zoneinfo](https://docs.python.org/3/library/zoneinfo.html) | Registro de data/hora no fuso horário de São Paulo |

---

## 📦 Pré-requisitos

- Python **3.10 ou superior** (necessário para `zoneinfo`)
- pip (gerenciador de pacotes Python)

---

## 🚀 Instalação e Execução

### 1. Clone o repositório

```bash
git clone https://github.com/seu-usuario/ecohive.git
cd ecohive
```

### 2. Crie e ative um ambiente virtual (recomendado)

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS / Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. Instale as dependências

```bash
pip install -r requirements.txt
```

### 4. Adicione os arquivos de imagem

Certifique-se de que os seguintes arquivos estão na raiz do projeto:

```
EcoHive_simbolo.png   # Imagem de fundo
EcoHive_1.png         # Logo do aplicativo
```

> O aplicativo funciona sem as imagens, utilizando cor de fundo como fallback.

### 5. Execute o aplicativo

```bash
python ecohive.py
```

---

## 📁 Estrutura do Projeto

```
ecohive/
├── ecohive.py            # Código-fonte principal
├── ecohive_.xlsx         # Banco de dados local (gerado automaticamente)
├── EcoHive_simbolo.png   # Imagem de fundo
├── EcoHive_1.png         # Logo
├── requirements.txt      # Dependências do projeto
└── README.md             # Este arquivo
```

---

## 🏗️ Arquitetura

O projeto segue uma arquitetura orientada a objetos com **duas camadas** separadas em classes:

### `BackEnd`
Responsável por toda a lógica de negócio e persistência de dados:

- `criar_excel()` — inicializa o arquivo `.xlsx` se não existir
- `ler_usuarios_df()` / `gravar_usuarios_df()` — CRUD da aba de usuários
- `inserir_usuario_excel()` — cadastro de novo morador com ID auto-incrementado
- `ler_descarte_df()` / `gravar_descarte()` — registro de descartes com data/hora
- `somatorio_por_tipo_por_bloco()` — agregação de kg reciclados por tipo, filtrado por bloco
- `salvar_pontuacao_usuario()` — persistência da pontuação acumulada

### `EcoHiveApp`
Herda de `ctk.CTk` (janela principal) e de `BackEnd`, sendo responsável pela interface gráfica:

- Gerencia **4 telas/frames**: Login, Cadastro, Home e Resgate
- Controla a lógica de navegação entre telas
- Lida com eventos de usuário (cliques, inputs, validações)
- Atualiza os elementos visuais de forma dinâmica

---

## 📖 Como Usar

### Primeiro acesso
1. Clique em **"Cadastre-se aqui"**
2. Preencha: usuário, e-mail, bloco, apartamento e senha
3. Clique em **"Cadastrar"** — você será redirecionado ao Login

### Registrar um descarte
1. Faça login com seu usuário ou e-mail e senha
2. Na tela Home, selecione o **tipo de resíduo** no ComboBox
3. Informe o **peso em kg** (ex: `1.5`)
4. Clique em **"Descartar"** — os pontos são creditados automaticamente

### Resgatar uma recompensa
1. Na tela Home, clique em **"Resgate"**
2. Veja as recompensas disponíveis e seus custos em pontos
3. Clique em **"Resgatar"** — os pontos são debitados automaticamente

### Tabela de pontuação

| Tipo de Resíduo | Pontos por kg |
|---|---|
| Papel | 5 |
| Vidro | 3 |
| Plástico | 10 |
| Metal | 15 |
| Eletrônico | 30 |
| Lata | 50 |

---

## 🗄️ Armazenamento de Dados

Os dados são persistidos localmente em `ecohive_.xlsx`, com duas abas:

**Aba `usuarios`**

| Campo | Tipo | Descrição |
|---|---|---|
| id | int | Identificador único auto-incrementado |
| usuario | str | Nome de usuário |
| bloco | str | Bloco do condomínio (A–D) |
| apartamento | int | Número do apartamento |
| email | str | E-mail do morador |
| senha | str | Senha (texto simples) |
| pontuacao | float | Pontuação acumulada |

**Aba `descarte`**

| Campo | Tipo | Descrição |
|---|---|---|
| data | str | Data do descarte (YYYY-MM-DD) |
| horario | str | Horário (HH:MM:SS) |
| usuario | str | Usuário responsável |
| bloco | str | Bloco do morador |
| apartamento | str | Apartamento do morador |
| tipo_lixo | str | Categoria do resíduo |
| peso_kg | float | Peso em kg |

> ⚠️ **Nota de segurança:** As senhas são armazenadas em texto simples. Em um contexto de produção, seria necessário aplicar hashing (ex: `bcrypt`).

---

## 🤝 Contribuição

Este projeto foi desenvolvido para fins acadêmicos. Sugestões e melhorias são bem-vindas:

1. Faça um **fork** do repositório
2. Crie uma branch para sua feature: `git checkout -b feature/minha-melhoria`
3. Commit suas mudanças: `git commit -m 'feat: adiciona minha melhoria'`
4. Push para a branch: `git push origin feature/minha-melhoria`
5. Abra um **Pull Request**

---

## 👤 Autor

**Caio Martins**
Estudante de Ciência da Computação — UNIP

- GitHub: Caio-Mar7ins [https://github.com/Caio-Mar7ins]
- LinkedIn: Caio Martins [https://www.linkedin.com/in/caio-martins-fernandes/]

---

*Desenvolvido como projeto acadêmico — UNIP, 2025.*

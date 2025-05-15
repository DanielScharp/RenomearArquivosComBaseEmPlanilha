# 📂 Renomear Arquivos com Base em Planilha (.xlsx)

Essa aplicação foi criada para **automatizar o processo de renomear arquivos (exemplo: imagens)** com base em uma lista presente em uma planilha do Excel. Ideal para rotinas de escritórios que lidam com a digitalização em massa de documentos.

## 🚀 Sobre o Projeto

Imagine que você digitalizou centenas de documentos em sequência. Os arquivos de imagem ficam com nomes como:

```
image001.jpg
image002.jpg
image003.jpg
...
```

E você possui uma planilha do Excel com os nomes reais que essas imagens deveriam ter:

```
ABC-1234
CBA-4321
AAA-1111
...
```

Essa aplicação faz a **mágica de renomear os arquivos automaticamente** com base nesses nomes, economizando tempo e evitando erros manuais.

---

## 🗂️ Estrutura Esperada

A aplicação busca os arquivos em uma **pasta específica na área de trabalho**:

```
📁 Área de Trabalho
└── 📁 Renomear Arquivos
     ├── 📁 arquivos
     │   ├── image001.jpg
     │   ├── image002.jpg
     │   └── ...
     ├── 📄 renomear.xlsx
     ├── 📁 bin
     │   └── RenomearArquivos.exe
     └── 🔗 Atalho para RenomearArquivos.exe
```

- **renomear.xlsx**: Planilha contendo os nomes (um por linha na **coluna A**).
- **arquivos/**: Pasta com os arquivos a serem renomeadas.
- **bin/**: Pasta onde fica o executável do programa.
- **Atalho (opcional)**: Um atalho do executável na raiz da pasta para facilitar o uso.
---

## ✅ Como Usar

1. **Preparar a estrutura de arquivos:**
   - Crie a pasta `Renomear Arquivos` na sua **Área de Trabalho**.
   - Dentro dela:
     - Crie uma subpasta `arquivos` e adicione os arquivos digitalizados.
     - Crie ou salve a planilha `renomear.xlsx` com os nomes na **coluna A**.
     - Copie o executável compilado (`RenomearArquivos.exe`) para a pasta `bin`.
    
2. **Criar um atalho (recomendado):**
   - Crie um **atalho do executável** (`RenomearArquivos.exe`) que está em `bin/` e posicione o atalho na **raiz da pasta `Renomear Arquivos`**.
   - Assim, o funcionário pode apenas clicar no atalho e rodar o processo.

3. **Executar a aplicação:**
   - Dê dois cliques no atalho ou no executável direto.
   - O programa irá:
     - Ler os nomes da planilha.
     - Identificar os arquivos no padrão `image*.*`.
     - Renomear os arquivos com os nomes da planilha na mesma ordem.

4. **Resultado:**
   - O programa renomeará os arquivos na pasta `arquivos` com os nomes da planilha, preservando a ordem.

> ⚠️ Se a quantidade de nomes e arquivos for diferente, o programa emitirá um aviso e pedirá confirmação antes de continuar.

---

## 📦 Publicando o Projeto

Para entregar essa aplicação para uso final (por exemplo, para um funcionário do escritório), siga este passo a passo:

1. Compile o projeto no modo **Release**.
2. Copie o executável (`.exe`) e todas as dependências da pasta `bin\Release\netX` para uma nova pasta chamada `bin`.
4. Crie um **atalho do `.exe`** (botão direito > "Criar atalho") e mova-o para a raiz da pasta.
5. Pronto! A aplicação está publicada e pronta para uso com um simples clique no atalho.


---


## 📎 Exemplo Visual

**Antes:**
```
image001.jpg
image002.jpg
image003.jpg
...
```

**Planilha:**
```
ABC-1234
CBA-4321
AAA-1111
...
```

**Depois:**
```
ABC-1234.jpg
CBA-4321.jpg
AAA-1111.jpg
...
```

---
## 🛠️ Tecnologias Utilizadas

- C# (.NET Console App)
- [EPPlus](https://github.com/EPPlusSoftware/EPPlus) (leitura de arquivos `.xlsx`)

---

## 🤝 Contribuições

Fique à vontade para sugerir melhorias ou enviar pull requests. Este projeto é simples, mas útil — e pode ser melhorado com sua ajuda!

---

## 📄 Licença

Uso **não-comercial** — este projeto utiliza EPPlus sob licença [Polyform Noncommercial](https://polyformproject.org/licenses/noncommercial/1.0.0/).

---

## ✉️ Contato

Se tiver dúvidas ou quiser relatar um problema, abra uma *issue* ou entre em contato!

---

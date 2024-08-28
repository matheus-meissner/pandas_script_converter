# Conversor e Formatação de Arquivos Excel com Python

Este projeto foi desenvolvido para automatizar o processamento, formatação e ajuste de arquivos Excel gerados a partir de dados extraídos de PDFs. Utilizando bibliotecas poderosas como `pandas` e `openpyxl`, o script oferece diversas funcionalidades para manipulação e formatação de planilhas Excel, garantindo uma apresentação de dados clara e padronizada.

## Funcionalidades

1. **Extração e Conversão de Dados:**
   - Extrai texto de arquivos PDF e converte os dados em um DataFrame utilizando `pandas`.
   - Remove linhas desnecessárias e limpa os dados para gerar uma planilha organizada.

2. **Formatação de Fontes:**
   - Padroniza a fonte e o tamanho de todas as células de uma planilha Excel.
   - Mantém textos de células específicas (como A1 e B1) em negrito, para destacá-las.

3. **Ajuste de Largura de Colunas:**
   - Ajusta a largura das colunas "A" e "B" para tamanhos especificados, garantindo uma visualização adequada dos dados.

4. **Aplicação de Bordas:**
   - Aplica bordas finas a todas as células que contêm texto na planilha Excel, melhorando a apresentação visual e separação dos dados.

5. **Quebra de Texto em Células:**
   - Ativa a quebra de linha para todas as células que contêm texto, exceto as células A1 e B1, preservando seu alinhamento original.

6. **Renomeação de Abas:**
   - Redefine o nome da aba da planilha para um título específico, como `IDJ_S1_E1`, garantindo uma melhor organização e identificação dos dados.

## Dependências

Certifique-se de ter as seguintes bibliotecas Python instaladas no seu ambiente:

- `pandas`
- `openpyxl`
- `pdfplumber`

Você pode instalá-las usando pip:

```bash
pip install pandas openpyxl pdfplumber
```

# Como usar:

```
## 1. Salve o DataFrame limpo no arquivo Excel
df_cleaned.to_excel(output_path, index=False, engine='openpyxl')
```
```
## 2. Padronize a fonte para Arial e tamanho 12 (ou qualquer outra fonte e tamanho que desejar)
standardize_font(output_path, font_name='Arial', font_size=12)
```
```
## 3. Ajuste as larguras das colunas A e B (A: 15 (aprox. 4cm), B: 64 (aprox. 16cm))
adjust_column_width(output_path, {'A': 20, 'B': 90})
```
```
## 4. Aplique bordas a todas as células que contêm texto
apply_borders(output_path)
```
```
## 5. Ative a quebra de texto em todas as células que contêm texto, exceto A1 e B1
wrap_text_in_cells(output_path)
```
```
## 6. Renomeie a aba ativa para "IDJ_S1_E1" (Renomeie como quiser)
rename_sheet(output_path, "IDJ_S1_E1")
```

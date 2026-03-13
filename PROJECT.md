# MD to DOCX Converter by Jair Lima

Conversor de arquivos Markdown (.md) para Word (.docx) com formatação completa e fiel ao padrão DOCX. Funciona como executável standalone chamável de qualquer pasta via terminal.

## Stack e Dependências

- **Python 3.14+**
- **python-docx 1.2+** — geração do arquivo DOCX
- **mistune 3.2+** — parser de Markdown (AST-based)
- **lxml 6+** — dependência do python-docx
- **PyInstaller 6+** — geração do executável .exe

## Estrutura de Arquivos

```
ConversorMD2DocX/
├── md2docx.py          # Script principal (conversor + CLI)
├── build.bat           # Script de build para gerar o .exe
├── requirements.txt    # Dependências Python
├── test_sample.md      # Arquivo MD de teste com todos os elementos
├── test_sample.docx    # DOCX gerado pelo teste
├── dist/
│   └── md2docx.exe     # Executável final (standalone)
└── PROJECT.md          # Este arquivo
```

## Comandos Essenciais

```bash
# Instalar dependências
pip install -r requirements.txt

# Executar via Python diretamente
python md2docx.py                     # Converter todos .md da pasta atual
python md2docx.py arquivo.md          # Converter arquivo específico
python md2docx.py arquivo.md --force  # Forçar sobrescrita

# Build do executável
build.bat                             # Gera dist/md2docx.exe

# Usar o executável (de qualquer pasta)
md2docx                               # Converter todos .md da pasta atual
md2docx README.md                     # Converter arquivo específico
md2docx README.md -f                  # Forçar sobrescrita
md2docx --folder C:\docs              # Pasta específica
md2docx --output C:\saida README.md   # Pasta de saída customizada
```

## Instalar o executável globalmente

Copiar `dist/md2docx.exe` para uma pasta no PATH do sistema:
- `C:\Windows\System32` (requer admin)
- `C:\Tools` ou similar (adicionar ao PATH manualmente)

## Elementos Markdown Suportados

| Elemento          | Suporte    | Estilo DOCX              |
|-------------------|-----------|--------------------------|
| H1 – H6           | ✅ completo | Heading 1-6             |
| Parágrafo         | ✅          | Normal                  |
| **Negrito**       | ✅          | Run bold                |
| *Itálico*         | ✅          | Run italic              |
| ***Bold+Italic*** | ✅          | Run bold+italic         |
| ~~Tachado~~       | ✅          | Run strikethrough       |
| `Código inline`   | ✅          | Courier New + highlight |
| Bloco de código   | ✅          | Code Block + shading    |
| Lista bullets     | ✅ aninhada | List Bullet 1-3         |
| Lista numerada    | ✅ aninhada | List Number 1-3         |
| Tabela            | ✅ alinhada | Word Table com borda    |
| Blockquote        | ✅          | Block Quote + borda esq |
| Link              | ✅          | Hyperlink DOCX nativo   |
| Imagem local      | ✅          | Inline picture          |
| Linha horizontal  | ✅          | Paragraph border        |
| YAML front matter | ✅ ignorado | —                       |

## Decisões Arquiteturais

- **mistune 3 (AST mode)**: Usado pela API `BaseRenderer` que recebe tokens estruturados em vez de HTML string, permitindo renderização direta para DOCX sem intermediário HTML.
- **Estilos nomeados DOCX**: Headings, Code Block, Block Quote usam estilos Word nomeados para compatibilidade máxima com editores DOCX.
- **PyInstaller --onefile**: Executável único, sem dependências externas após build.
- **Encoding UTF-8**: `sys.stdout.reconfigure(encoding='utf-8')` resolve problema com terminal Windows (cp1252).

## Estado Atual

- ✅ Conversor funcional e testado
- ✅ Executável `dist/md2docx.exe` gerado
- ✅ Todos os elementos comuns suportados

## Próximos Passos (sugestões)

- Suporte a footnotes DOCX nativo (atualmente ignorado)
- Task lists com checkboxes (☑ / ☐)
- Ícone customizado no .exe
- Suporte a múltiplas pastas recursivas (`--recursive`)
- Logging para arquivo ao invés de stdout

## Problemas Conhecidos

- Imagens referenciadas por URL remota não são baixadas (exibe texto de alt)
- Tabelas muito largas podem extrapolar a margem (sem ajuste automático de coluna)

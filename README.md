# RasCol Automation

Extrai o **Relatório de Pontos de Operação** do sistema [RasCol](https://rascol.rassystem.com.br) para a filial Jaboatão (empresa Locar), baixando um arquivo `.xls` por veículo para cada janela de 7 dias, e opcionalmente convertendo os dados em shapefiles de segmentos GPS prontos para entrega.

---

## Requisitos

- Python 3.10+
- Dependências declaradas no `pyproject.toml` / `requirements.txt` do repositório pai (`Locar/`)
- Pacote irmão `inlog_automation` presente na mesma pasta `Locar/`
- Google Chrome instalado (o WebDriver é gerenciado automaticamente)

---

## Configuração

As credenciais e a filial são lidas da seção `[RASCOL]` do arquivo `Locar/dependencias/.env`:

```ini
[RASCOL]
usuario = 
senha   = 
filial  = 
```

Os diretórios de trabalho criados automaticamente dentro de `Locar/dependencias/`:

| Pasta | Conteúdo |
|---|---|
| `rascol_downloads/` | Arquivos `.xls` baixados do RasCol |
| `rascol_shapes/` | ZIPs de shapefiles gerados (quando não há ZIP da Inlog para o dia) |
| `shapes/` | ZIPs da Inlog — shapefiles RasCol são adicionados aqui quando o ZIP do dia já existe |

---

## Execução

### Modo desenvolvimento

```bash
# a partir de Locar/
python rascol_automation/run.py
```

### Executável compilado

```bash
# a partir de Locar/rascol_automation/
pyinstaller RasColAutomation.spec
```

Copie `dist/RasColAutomation.exe` para a pasta `Locar/` (ao lado de `InlogAutomation.exe`).

---

## Fluxo de extração

1. Abre Chrome → login CAS → seleciona empresa Locar → filial Jaboatão → Relatório de Pontos de Operação
2. Define Data/Hora da primeira janela (**D_início 05:00 → D_fim+1 05:00**)
3. Seleciona rótulo **DOMICILIAR** → aguarda lista de veículos carregar
4. Para cada janela de até 7 dias:
   - Atualiza Data/Hora (somente a partir da segunda janela)
   - Para cada veículo: seleciona → Pesquisar → Exportar (se houver resultado) → aguarda download
5. Opcional: converte os `.xls` em shapefiles de segmentos e os empacota em ZIPs por dia

---

## Geração de Shapefiles

Ativado pelo checkbox **"Gerar Shapefiles"** na GUI.

- Lê placa (célula B4) e contrato (E3) do cabeçalho de cada Excel
- Dados GPS a partir da linha 7: `data/hora`, `latitude`, `longitude`, `velocidade`
- Agrupa por **dia operacional** = `(datahora − 5 h).date`
- Dias além da data final selecionada são descartados (evita artefatos de borda)
- Cada segmento GPS consecutivo vira uma feição `LineString` (atributos: `DATAHORA`, `VELOCIDADE`)
- Destino do ZIP por dia:
  - Se `dependencias/shapes/Shapes - Jaboatao - DD.MM.YYYY.zip` já existir → acrescenta ali
  - Caso contrário → cria/acrescenta em `dependencias/rascol_shapes/`
- Excels processados são apagados após a geração

---

## Estrutura do pacote

```
rascol_automation/
├── run.py                        # entry-point
├── config/
│   ├── settings.py               # caminhos (RASCOL_DOWNLOAD_DIR, RASCOL_SHAPES_DIR, INLOG_SHAPES_DIR)
│   └── rascol_config.py          # lê [RASCOL] do .env
├── core/
│   ├── browser.py                # abre Chrome com prefs corretas
│   ├── waits.py                  # wait_load_rascol + re-exports de inlog_automation
│   └── auth.py                   # login CAS, seleção empresa/filial, navegação ao relatório
├── extractors/
│   └── extractor_pontos.py       # PontosExtractor — loop janelas × veículos
├── processors/
│   └── processor_shapes.py       # ShapesProcessor — Excel → LineString shapefiles → ZIP
└── gui/
    ├── main_gui.py               # RasColGUI (calendário, credenciais, checkbox shapes)
    └── runner.py                 # integra GUI ↔ PontosExtractor ↔ ProgressWindow
```

# Aspen Plus V15 MCP Server

MCP (Model Context Protocol) server that lets Claude — in Claude Code, Claude Desktop or any MCP
client — drive **Aspen Plus V15** through its COM interface: open and run simulations, read and
write any node of the tree, edit the flowsheet, diagnose convergence and, what no other public Aspen
MCP server does, **design RadFrac Column Internals** (tray and packed sections, hydraulics, tray
efficiency and column height).

By **Eduardo Mariola** — Chemical Engineering, UNESP (Brazil).

## Why this exists

I spent two semesters designing a benzyl alcohol plant (hydrolysis of benzyl acetate): patent
search and route selection, the Aspen Plus flowsheet I built block by block — reactor, five
distillation columns, decanter and recycles —, economic evaluation, P&IDs in AutoCAD Plant 3D, and
now the 3D model, HAZOP and LOPA.

When the time came to model the plant in 3D, the tray columns came out absurd: under 1 m in
diameter and 19 to 27 m tall. Without Column Internals, the economic evaluator assumes 24-inch tray
spacing and its own tray efficiency for every column. None of the public Aspen MCP servers could
touch Column Internals, so I built that here, in Claude Code sessions on my own machine, on top of
a year of project context.

Result on my process: tray internals defined for four columns through MCP calls, the file re-opened
from scratch and run with `PER_ERROR = 0`, product at 99.84 wt% benzyl alcohol, and column heights of
2.8 to 14.5 m with realistic L/D.

> **Em português.** Passei dois semestres projetando uma planta de álcool benzílico (busca de
> patentes, rota, fluxograma no Aspen montado bloco a bloco, avaliação econômica, P&IDs no Plant 3D
> e agora o modelo 3D, HAZOP e LOPA). Na hora do 3D, as torres de pratos saíam com menos de 1 m de
> diâmetro e 19 a 27 m de altura, porque sem internos de coluna o avaliador econômico usa
> espaçamento de 24" para tudo. Nenhum servidor MCP público do Aspen mexia em internos, então
> construí isso aqui, no Claude Code e na minha máquina, com o contexto de um ano de projeto.

## Features

- **Column Internals (new)**: create `INT-1`, add tray or packed sections over stage ranges; set
  tray type, passes, spacing, diameter, deck gauge, weir height, hole size, downcomer clearance and
  overall efficiency; interactive sizing or rating; read flooding, downcomer backup and pressure
  drop; O'Connell efficiency from the converged column; tangent-to-tangent height and L/D
- **Process simulation control**: open (`.apw` via `InitFromFile2`, `.bkp` via `InitFromArchive2`),
  run, reinitialize, save, stop
- **Work in your own window**: `attach` binds to a simulation already open in the Aspen GUI
- **Robust on modest machines**: long runs are started asynchronously and polled; when DCOM cannot
  start Aspen (low memory), the server launches `AspenPlus.exe` and attaches
- **Data access**: read and write any node by path; streams, blocks, components, reactions, utilities
- **Flowsheet editing**: add/remove blocks and streams, connect ports, RadFrac stages, feeds, products
- **Diagnostics**: convergence report after every run, knowledge base, incomplete-input scan, tear streams

## Requirements

- Windows (Aspen Plus is driven through COM)
- Aspen Plus V15 (GUI version 41.0)
- Python 3.10 or higher
- `pywin32` and `fastmcp` (installed with the package)

## Installation

1. Clone this repository:
```bash
git clone https://github.com/Eduardo-Mariola/aspen-plus-mcp.git
cd aspen-plus-mcp
```

2. Create a virtual environment:
```bash
python -m venv .venv
.venv\Scripts\activate
```

3. Install the package:
```bash
pip install -e .
```

## Configuration

### Claude Code

```bash
claude mcp add --scope user aspen-plus -- "C:\path\to\aspen-plus-mcp\.venv\Scripts\python.exe" -m aspen_mcp.server
```

The tools show up as `mcp__aspen-plus__*` in the next session.

### Claude Desktop

Add to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "aspen-plus": {
      "command": "C:\\path\\to\\aspen-plus-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-m", "aspen_mcp.server"]
    }
  }
}
```

**Important notes:**
- Use the full absolute path to the virtual environment's Python
- Restart Claude Desktop completely after editing the file
- If Aspen Plus is installed elsewhere, set `ASPEN_PLUS_EXE` to the path of `AspenPlus.exe`

## Usage

### Running the server

```bash
python -m aspen_mcp.server
```

### Available tools

#### Simulation
- `open_file` - Open `.apw`/`.apwz`/`.bkp` (attaches if the file is already open in the GUI)
- `attach` - Work on a simulation already open in the Aspen Plus window
- `run` / `reinit_and_run` - Run (asynchronously, polled) and return the convergence report
- `save` - Save, or save as a new path
- `status`, `probe`, `stop_simulation`, `reinit`, `close_file`, `visible`, `new_simulation`

#### Column Internals (new)
- `add_column_internals` - Create a configuration (`INT-1`) on a RadFrac and make it active
- `add_internals_section` / `remove_internals_section` - Tray or packed section over a stage range
- `configure_tray_section` - Tray type, passes, spacing, mode (sizing/rating), diameter, deck gauge,
  weir height, hole size, downcomer clearance, overall efficiency, foaming factor
- `configure_packed_section` - Packing type, vendor, material, size, packed height or HETP
- `set_internals_param` - Any `CA_xxx` input, validated against Aspen's option list
- `describe_internals_params` - Aspen's own description, unit and options of each `CA_xxx` input
- `list_column_internals` - Sections, geometry and results
- `get_column_hydraulics` - Flooding, downcomer backup, pressure drop, per section or per stage
- `oconnell_efficiency` - Overall tray efficiency from the converged column
- `column_height_estimate` - Real trays, tangent-to-tangent height and L/D
- `resize_column_stages` - Change `NSTAGE` keeping feeds, bottoms draw and sections consistent

#### Columns
- `set_column_stages`, `set_feed_stage`, `set_product_stage`, `set_column_pressure`
- `set_condenser_type`, `set_reboiler_type`, `add_side_duty`, `remove_side_duty`

#### Flowsheet, streams and blocks
- `list_all_blocks`, `get_block`, `add_block`, `remove_block`, `set_param`, `block_status`
- `list_all_streams`, `get_stream`, `add_stream`, `remove_stream`, `set_stream_param`,
  `set_stream_composition`, `set_stream_composition_batch`
- `connect`, `connect_port`, `disconnect`, `list_block_ports`, `list_all_ports`, `flowsheet_topology`,
  `configure_fsplit`

#### Components, properties, reactions, utilities
- `list_components`, `add_component`, `remove_component`, `get_property_method`, `set_property_method`
- `add_reaction_set`, `add_reaction`, `remove_reaction`, `remove_reaction_set`, `list_reaction_sets`,
  `setup_block_reaction`
- `add_utility`, `batch_add_utilities`, `list_utilities`, `get_utility`, `remove_utility`

#### Paths, analysis and diagnostics
- `get_value`, `set_value`, `insert_row`, `set_label`, `explore`
- `sensitivity`, `sensitivity_advanced`, `diagnose`, `search_convergence_knowledge`,
  `find_incomplete_inputs`, `validate_block`, `simulation_warnings`, `list_tear_streams`,
  `set_tear_estimate`, `export_report_file`, `help`

### Example node paths

```
# Stream temperature and component mass fraction
\Data\Streams\S1\Output\TEMP_OUT\MIXED
\Data\Streams\S1\Output\MASSFRAC\MIXED\WATER

# RadFrac input and stage profile
\Data\Blocks\T-102\Input\NSTAGE
\Data\Blocks\T-102\Input\FEED_STAGE\27
\Data\Blocks\T-102\Output\B_TEMP\5

# Column Internals input: <param>\<config>\<section>
\Data\Blocks\T-102\Input\CA_TRAY_SPC\INT-1\CS-1
\Data\Blocks\T-102\Input\CA_DIAM\INT-1\CS-1
\Data\Blocks\T-102\Input\CA_EFF\INT-1\CS-1

# Column Internals results (tray, rating mode)
\Data\Blocks\T-102\Output\CA_FLD_FAC5\INT-1\CS-1       # max % flood
\Data\Blocks\T-102\Output\CA_DCBSPACE1\INT-1\CS-1      # max downcomer backup, % of spacing
\Data\Blocks\T-102\Output\CA_DPSEC3\INT-1\CS-1         # section pressure drop

# Run status
\Data\Results Summary\Run-Status\Output\PER_ERROR
```

## Usage examples

### Attach to the open simulation and run
```python
attach(file_path="C:/projects/plant.apw")
reinit_and_run()
get_value(path="\\Data\\Streams\\PRODUCT\\Output\\MASSFRAC\\MIXED\\BZOH")
```

### Size a tray column, then rate it with a standard diameter
```python
add_column_internals(block_name="T-103")
add_internals_section(block_name="T-103", section="CS-1", stage_start=2, stage_end=29)
configure_tray_section(block_name="T-103", section="CS-1", tray_type="SIEVE",
                       tray_spacing_m=0.30, mode="INTER-SIZING",
                       deck_thickness="14 GAUGE", weir_height_mm=50)
reinit_and_run()
get_column_hydraulics(block_name="T-103")                    # sized diameter (CA_DIAM6)
oconnell_efficiency(block_name="T-103", light_key="BZOH", heavy_key="BZOAC")

configure_tray_section(block_name="T-103", section="CS-1", mode="RATING",
                       diameter_m=0.60, overall_efficiency=0.72, tray_spacing_m=0.30)
reinit_and_run()
get_column_hydraulics(block_name="T-103", per_stage=True)    # % flood and downcomer backup per tray
column_height_estimate(block_name="T-103")                   # real trays, T-T height, L/D
save(file_path="C:/projects/plant - internals.apw")
```

### Packed section
```python
add_internals_section(block_name="T-101", section="CS-1", stage_start=2, stage_end=14,
                      internal_type="PACKING")
configure_packed_section(block_name="T-101", section="CS-1", packing_type="MELLAPAK",
                         vendor="SULZER", material="STANDARD", size="250Y", hetp_m=0.50)
```

## Lessons from Aspen Plus V15

- An `.apw` must be opened with `InitFromFile2`; archive-loading it leaves `\Data\Blocks` empty.
- When DCOM cannot start `aspenplus.exe -Automation` (the APED LocalDB fails with error 575 on a
  machine short of memory), opening `AspenPlus.exe` directly and attaching through the Running
  Object Table works.
- A run that returns no results at all: read the `.his`. An incomplete property analysis (a Txy with
  `RANGE` selected but only a pressure list filled) stops input translation for the whole flowsheet.
- Tray results use different keys from packing, and from sizing to rating. Aspen's section height
  and pressure drop count one tray per theoretical stage and ignore `CA_EFF`, so
  `column_height_estimate` applies the efficiency itself.
- Three-phase RadFrac blocks leave `B_K` and `LIQ_FLOW_MS` empty; the tools fall back to y/x and
  `HYD_LMF`.
- Feeding the hydraulic pressure drop back into the column raises bottoms temperatures and can push
  a reboiler above its steam level (`TEMPERATURES CROSS WITH UTILITY` in the `.his`).

## Development

```bash
pip install -e .
python -c "import aspen_mcp.server"        # import check
```

Project layout:

```
src/aspen_mcp/
  server.py            MCP tool registration (FastMCP)
  com_bridge.py        COM apartment thread, open/attach, async run
  tools/internals.py   Column Internals tools
  tools/*.py           simulation, streams, blocks, columns, analysis, reactions, utilities
  knowledge/           convergence knowledge base
docs/blocks/           parameter notes per Aspen block type
```

## Credits

The COM bridge and the general-purpose tools are adapted from the MIT-licensed
[aspen-mcp](https://github.com/thirakorn-mokkawes-59/aspen-mcp); its copyright notice is kept in
[LICENSE](LICENSE).

## License

MIT

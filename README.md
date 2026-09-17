# Global Markets Case Exploration

Python analysis of client revenue, FX margins, product contributions, sector margins and CNY trade volumes for 2020–2022. Results are exported to an Excel workbook, including a monthly revenue chart and notes and assumptions.

## Run in your browser with GitHub Codespaces

### 1. Open your own Codespace

1. Sign in to your GitHub account.
2. Open [this project in GitHub Codespaces](https://codespaces.new/inc0gn1k0/gm-case-exploration).
3. Choose the `main` branch if prompted, create the Codespace and wait for the browser editor to open.

This creates your own environment. You do not need access to someone else's running Codespace. Codespaces usage is subject to your account's allowance and billing settings.

### 2. Install the editor extensions

Open **Extensions** in the left sidebar. Search for and install these extensions, both published by **Microsoft**:

| Extension | Exact extension ID | Purpose |
| --- | --- | --- |
| Python | `ms-python.python` | Python interpreter selection and editor support |
| Jupyter | `ms-toolsai.jupyter` | **Run Cell**, **Run Above** and **Run Below** controls for `# %%` cells, plus the Interactive Window |

If prompted, choose **Install in Codespaces**. If an extension is already installed, ensure it is enabled in this Codespace.

Installing Python packages does not install these editor extensions.

### 3. Install the Python dependencies

Open **Terminal → New Terminal**. From the project root, run:

```bash
python -m pip install -r requirements.txt
```

The requirements include `ipykernel`, which lets Jupyter execute Python cells, as well as the analysis and Excel-export libraries.

If you are returning to an existing Codespace and want the latest published version, first run `git pull`, then repeat the installation command. If Git reports conflicting local edits, resolve those before continuing.

### 4. Select Python and reload the editor

1. Open the Command Palette using **F1** or **View → Command Palette**.
2. Run **Python: Select Interpreter** and choose the environment where you installed the requirements.
3. Run **Developer: Reload Window** from the Command Palette.
4. Open `gm-exploration-cleaned.py`.

If you are unsure which Python environment the terminal uses, run:

```bash
python -c "import sys; print(sys.executable)"
```

Select that interpreter path. When the Interactive Window asks for a kernel, select the same Python environment.

### 5. Run the analysis

**Option A: run individual cells**

The script is divided into cells marked by `# %%`.

1. Click **Run Cell** above the first **SETUP** cell.
2. Run the **Query 1** through **Query 8** cells in order using **Run Cell**.
3. Read the printed tables and chart in the Python Interactive Window.

To execute everything together, click **Run Below** above the first **SETUP** cell. This includes setup, so you do not need to run setup separately first.

**Important:** the setup cell recreates the output workbook. Rerunning it clears previous exports; run all query cells again afterward. Keep the same kernel/session while running the cells, because later cells use variables created earlier.

**Option B: run the whole script from the terminal**

```bash
python gm-exploration-cleaned.py
```

This executes setup and all eight queries. The editor extensions are only needed for the cell-based workflow. The chart is embedded in the Excel output even if the terminal run does not open a chart window.

### 6. Download the results

The script creates:

```text
GM-Presentation/GM-Query-Outputs-Cleaned.xlsx
```

Find the workbook in the Explorer, right-click it and select **Download**. Open the downloaded file in Excel or another compatible spreadsheet application. The workbook contains the query outputs, the monthly revenue chart, methodology, and **Notes & Assumptions**.

Generated results are ignored by Git. Close the output workbook before rerunning the script locally if Excel has locked the file.

## Troubleshooting cells

- **No Run Cell buttons:** check that Microsoft's Python and Jupyter extensions are installed and enabled in the Codespace. Open the `.py` file, confirm its language mode is Python, and run **Developer: Reload Window**. If necessary, enable **Editor: Code Lens** and **Jupyter: Enable Cell Code Lens** in Settings.
- **Kernel selection or missing `ipykernel`:** select the Python environment used for dependency installation. Rerun `python -m pip install -r requirements.txt` in that environment.
- **`ModuleNotFoundError`:** the selected kernel may differ from the terminal's Python. Compare the interpreter path using the command above, select the matching kernel and restart it.
- **Variables are undefined:** run the setup cell, then run the queries in order in the same Interactive Window.
- **`__file__` is undefined:** run the original `.py` file using its **Run Cell** controls or use the terminal command. Do not paste the setup code into a separate notebook or Python console; it uses the script's path to locate the data files.
- **Missing input files:** keep the repository's folder structure intact and run terminal commands from the project root.

## Project files

| File | Purpose |
| --- | --- |
| `gm-exploration-cleaned.py` | Main analysis, split into setup and eight query cells |
| `plot_theme.py` | Chart colors and formatting |
| `requirements.txt` | Python dependencies, including the interactive kernel |
| `GM-Case-study-Data-2026.xlsx` | Working input workbook read by the script |
| `Source-Material/GM Case study Data 2026.xlsx` | Original source workbook |
| `Source-Material/exchange-rates-zar-2020-2022.csv` | Monthly USD/ZAR conversion rates |

Query 7 describes a method for ranking exporter-shaped behavior; it intentionally does not calculate a ranking. See the exported methodology and notes for the definitions and assumptions behind the other results.

## Local setup

Clone the repository, open its folder in VS Code, install the same two extensions and follow the dependency and interpreter steps above. A virtual environment is recommended:

```bash
python -m venv .venv
```

Activate it using `source .venv/bin/activate` on macOS/Linux or `.venv\Scripts\Activate.ps1` in Windows PowerShell, then install the requirements and select `.venv` as your interpreter and kernel.

## Reference

- [VS Code Python Interactive Window](https://code.visualstudio.com/docs/python/jupyter-support-py)
- [VS Code Jupyter kernel selection](https://code.visualstudio.com/docs/datascience/jupyter-kernel-management)
- [GitHub Codespaces creation links](https://docs.github.com/en/codespaces/setting-up-your-project-for-codespaces/setting-up-your-repository/facilitating-quick-creation-and-resumption-of-codespaces)

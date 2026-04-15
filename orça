import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import LineChart, Reference
from weasyprint import HTML
import os

# --- 1. DATA PREPARATION ---
project_info = {
    "Proprietário": "MD Engenharia / Cliente Final",
    "Localização": "Anápolis, GO - Condomínio Fechado",
    "Área Total": "200 m²",
    "Padrão": "Alto",
    "Modalidade": "Aquisição de Terreno e Construção",
    "Base": "SINAPI GO (Desonerado) - Abr/2026",
    "BDI": "25.0%"
}

budget_items = [
    {"Item": "1.0", "Desc": "SERVIÇOS PRELIMINARES", "Código": "99059", "Unid": "un", "Qtd": 1, "Unit": 15000.00},
    {"Item": "2.0", "Desc": "INFRAESTRUTURA (Sapatas e Baldrames)", "Código": "96522", "Unid": "m³", "Qtd": 22, "Unit": 3800.00},
    {"Item": "3.0", "Desc": "SUPRAESTRUTURA (Pilares/Vigas/Lajes)", "Código": "103671", "Unid": "m³", "Qtd": 35, "Unit": 4200.00},
    {"Item": "4.0", "Desc": "PAREDES E PAINÉIS (Vedação)", "Código": "87503", "Unid": "m²", "Qtd": 480, "Unit": 125.00},
    {"Item": "5.0", "Desc": "COBERTURA (Telha Termoacústica)", "Código": "92541", "Unid": "m²", "Qtd": 210, "Unit": 220.00},
    {"Item": "6.0", "Desc": "INSTALAÇÕES HIDROSSANITÁRIAS", "Código": "95471", "Unid": "un", "Qtd": 1, "Unit": 35000.00},
    {"Item": "7.0", "Desc": "INSTALAÇÕES ELÉTRICAS/FOTOVOLTAICA", "Código": "91863", "Unid": "un", "Qtd": 1, "Unit": 48000.00},
    {"Item": "8.0", "Desc": "REVESTIMENTOS INTERNOS (Alto Padrão)", "Código": "87250", "Unid": "m²", "Qtd": 650, "Unit": 180.00},
    {"Item": "9.0", "Desc": "ESQUADRIAS DE ALUMÍNIO PRETO", "Código": "94570", "Unid": "m²", "Qtd": 45, "Unit": 1800.00},
    {"Item": "10.0", "Desc": "PISCINA EM CONCRETO ARMADO", "Código": "98147", "Unid": "un", "Qtd": 1, "Unit": 45000.00},
    {"Item": "11.0", "Desc": "SERVIÇOS COMPLEMENTARES/PAISAGISMO", "Código": "98511", "Unid": "un", "Qtd": 1, "Unit": 25000.00},
]

# --- 2. CREATE EXCEL (XLSX) ---
wb = Workbook()
ws_summary = wb.active
ws_summary.title = "Resumo do Projeto"

# Styling
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
white_font = Font(color="FFFFFF", bold=True)
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# Summary Tab
ws_summary.append(["DADOS DO PROJETO - BUILDFLOW SINAPI"])
ws_summary.merge_cells("A1:B1")
ws_summary["A1"].font = Font(size=14, bold=True)

for i, (k, v) in enumerate(project_info.items(), 2):
    ws_summary.cell(row=i, column=1, value=k).font = Font(bold=True)
    ws_summary.cell(row=i, column=2, value=v)

# Budget Tab
ws_budget = wb.create_sheet("Orçamento Analítico")
headers = ["Item", "Descrição", "Código SINAPI", "Unid", "Qtd", "Preço Unit (R$)", "Total (R$)"]
ws_budget.append(headers)

for cell in ws_budget[1]:
    cell.fill = header_fill
    cell.font = white_font

total_geral = 0
for idx, row in enumerate(budget_items, 2):
    total = row["Qtd"] * row["Unit"]
    total_geral += total
    ws_budget.append([row["Item"], row["Desc"], row["Código"], row["Unid"], row["Qtd"], row["Unit"], total])
    ws_budget.cell(row=idx, column=6).number_format = '#,##0.00'
    ws_budget.cell(row=idx, column=7).number_format = '#,##0.00'

# Schedule Tab (Simplified)
ws_sched = wb.create_sheet("Cronograma CEF")
months = [f"Mês {i}" for i in range(1, 13)]
ws_sched.append(["Etapa"] + months + ["Total"])
for cell in ws_sched[1]:
    cell.fill = header_fill
    cell.font = white_font

# Distribute values roughly
for row in budget_items:
    ws_sched.append([row["Desc"]] + ["8.33%"] * 12 + ["100%"])

wb.save("pci_orcamento_anapolis_v1.xlsx")

# --- 3. CREATE PDF ---
html_content = f"""
<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8">
<style>
    @page {{
        size: A4;
        margin: 20mm;
        background-color: #ffffff;
    }}
    body {{
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        color: #333;
        line-height: 1.6;
    }}
    .header {{
        border-bottom: 2px solid #1F4E78;
        margin-bottom: 30px;
        padding-bottom: 10px;
    }}
    h1 {{ color: #1F4E78; font-size: 22pt; margin: 0; }}
    h2 {{ color: #2E75B6; border-left: 5px solid #1F4E78; padding-left: 10px; margin-top: 30px; font-size: 16pt; }}
    .info-box {{ background: #f4f4f4; padding: 15px; border-radius: 5px; margin-bottom: 20px; }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
    th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; font-size: 10pt; }}
    th {{ background-color: #1F4E78; color: white; }}
    .chart-box {{ margin-top: 30px; text-align: center; border: 1px dashed #ccc; padding: 20px; }}
    .footer {{ position: fixed; bottom: 0; width: 100%; text-align: center; font-size: 8pt; color: #999; }}
</style>
</head>
<body>
    <div class="header">
        <h1>Relatório Técnico: BuildFlow SINAPI</h1>
        <p>Projeto: Residência Alto Padrão 200m² - Anápolis/GO</p>
    </div>

    <div class="info-box">
        <strong>Responsável Técnico:</strong> MD Engenharia<br>
        <strong>Data de Referência:</strong> 14 de Abril de 2026<br>
        <strong>Base de Custos:</strong> SINAPI GO - Desonerado
    </div>

    <h2>1. Memorial Descritivo de Alto Padrão</h2>
    <p>A residência será executada com foco em excelência construtiva. A infraestrutura em sapatas isoladas garantirá a estabilidade em solo goiano. Os revestimentos incluem porcelanatos de grandes formatos (90x90 ou superior) em todas as áreas sociais.</p>
    <ul>
        <li><strong>Esquadrias:</strong> Linha Gold em alumínio preto com vidros laminados.</li>
        <li><strong>Climatização:</strong> Infraestrutura para Split em todos os quartos e salas.</li>
        <li><strong>Lazer:</strong> Piscina em concreto armado com revestimento em pedra vulcânica (Hijau) e sistema de aquecimento solar dedicado.</li>
    </ul>

    <h2>2. Cronograma e Gráfico de Gantt (Visualização)</h2>
    <table>
        <tr><th>Etapa</th><th>M1</th><th>M2</th><th>M3</th><th>M4</th><th>M5</th><th>M6</th><th>M7</th><th>M8</th><th>M9</th><th>M10</th><th>M11</th><th>M12</th></tr>
        <tr><td>Fundações</td><td>■</td><td>■</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>
        <tr><td>Estrutura</td><td></td><td>■</td><td>■</td><td>■</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>
        <tr><td>Alvenaria</td><td></td><td></td><td></td><td>■</td><td>■</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>
        <tr><td>Instalações</td><td></td><td></td><td>■</td><td>■</td><td>■</td><td>■</td><td></td><td></td><td></td><td></td><td></td><td></td></tr>
        <tr><td>Acabamento</td><td></td><td></td><td></td><td></td><td></td><td></td><td>■</td><td>■</td><td>■</td><td>■</td><td>■</td><td></td></tr>
        <tr><td>Piscina</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>■</td><td>■</td><td></td><td></td></tr>
        <tr><td>Limpeza</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>■</td></tr>
    </table>

    <h2>3. Curva S de Desembolso (Estimada)</h2>
    <div class="chart-box">
        <svg width="500" height="200" viewBox="0 0 500 200">
            <polyline points="0,180 50,170 100,150 150,120 200,90 250,70 300,50 350,35 400,25 450,15 500,5" 
                      fill="none" stroke="#1F4E78" stroke-width="3" />
            <text x="0" y="195" font-size="10">Mês 1</text>
            <text x="450" y="195" font-size="10">Mês 12</text>
            <text x="10" y="20" font-size="10" transform="rotate(-90, 10, 20)">% Acumulado</text>
        </svg>
        <p><em>Curva S representa o investimento acumulado ao longo dos 12 meses.</em></p>
    </div>

    <div class="footer">Gerado automaticamente pelo sistema BuildFlow SINAPI - MD Engenharia</div>
</body>
</html>
"""

with open("memorial_temp.html", "w", encoding="utf-8") as f:
    f.write(html_content)

HTML(filename="memorial_temp.html").write_pdf("memorial_tecnico_v1.pdf")

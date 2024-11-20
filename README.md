from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Cria um workbook e uma planilha
wb = Workbook()
ws = wb.active
ws.title = "Checklist Cirurgia Segura"

# Estilos
header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill("solid", fgColor="4F81BD")
alignment = Alignment(horizontal="center", vertical="center")
thin_border = Border(
    left=Side(style="thin"), 
    right=Side(style="thin"), 
    top=Side(style="thin"), 
    bottom=Side(style="thin")
)

# Cabeçalho da Planilha
headers = [
    "Fase da Cirurgia",
    "Item",
    "Sim",
    "Não",
    "Observações"
]

# Adiciona cabeçalho
ws.append(headers)
for col, header in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=col)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = alignment
    cell.border = thin_border

# Adiciona as fases e os itens
data = [
    ["1. Antes da Indução Anestésica", "Identidade do paciente confirmada", "", "", ""],
    ["", "Procedimento e local cirúrgico marcados e confirmados", "", "", ""],
    ["", "Alergias conhecidas revisadas", "", "", ""],
    ["", "Equipamentos específicos disponíveis", "", "", ""],
    ["", "Planejamento para riscos específicos realizado", "", "", ""],
    ["2. Antes da Incisão Cirúrgica", "Todos os membros da equipe se apresentaram", "", "", ""],
    ["", "Procedimento e local confirmados novamente", "", "", ""],
    ["", "Instrumentos e equipamentos verificados", "", "", ""],
    ["", "Antibióticos profiláticos administrados (se aplicável)", "", "", ""],
    ["", "Discussão de riscos específicos para a cirurgia realizada", "", "", ""],
    ["3. Antes do Paciente Deixar a Sala", "Contagem final de instrumentos, compressas e agulhas realizada", "", "", ""],
    ["", "Procedimento finalizado conforme planejado", "", "", ""],
    ["", "Materiais coletados identificados corretamente", "", "", ""],
    ["", "Plano pós-operatório discutido com a equipe", "", "", ""],
    ["", "Comunicação feita com a equipe de recuperação", "", "", ""]
]

for row_data in data:
    ws.append(row_data)

# Ajusta larguras das colunas
column_widths = [25, 70, 10, 10, 30]
for i, width in enumerate(column_widths, start=1):
    ws.column_dimensions[chr(64 + i)].width = width

# Salva o arquivo
file_path = "/mnt/data/Checklist_Cirurgia_Segura.xlsx"
wb.save(file_path)
file_path

- 👋 Hi, I’m @RafaelaCalado
- 👀 I’m interested in ...
- 🌱 I’m currently learning ...
- 💞️ I’m looking to collaborate on ...
- 📫 How to reach me ...
- 😄 Pronouns: ...
- ⚡ Fun fact: ...

<!---
RafaelaCalado/RafaelaCalado is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->

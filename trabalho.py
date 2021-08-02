from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

data = {
    "João": {
        "Matemática": 65,
        "Geografia": 78,
        "Inglês": 87,
        "História": 98
    },
    "Maria": {
        "Matemática": 54,
        "Geografia": 73,
        "Inglês": 82,
        "História": 65
    },
    "Lucas":{
        "Matemática": 99,
        "Geografia": 99,
        "Inglês": 98,
        "História": 98
    },
    "Gisleine": {
        "Matemática": 78,
        "Geografia": 73,
        "Inglês": 56,
        "História": 88
    }
}


wb = Workbook()
ws = wb.active
ws.title = "Notas"

headings = ['Nome'] + list(data['João'].keys())
ws.append(headings)

for pessoa in data:
    notas = list(data[pessoa].values())
    ws.append([pessoa] + notas)


wb.save("NovoNotas.xlsx")
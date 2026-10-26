"""gspread FALSO para testes offline: planilha em memória, nenhuma rede, dados fictícios."""
import copy
import datetime as _dt
import re


class WorksheetNotFound(Exception):
    pass


def _num(letras):
    n = 0
    for c in letras:
        n = n * 26 + ord(c) - 64
    return n


class Worksheet:
    def __init__(self, title, cols=3):
        self.title, self.linhas, self.col_count = title, [], cols

    def get_all_values(self):
        return copy.deepcopy(self.linhas)

    def update(self, *args, range_name=None, values=None, **_):
        if args:
            range_name, values = args[0], args[1]
        ini = range_name.split(":")[0]
        lin = int(re.search(r"\d+", ini).group(0))
        col = _num(re.match(r"[A-Z]+", ini).group(0))
        while len(self.linhas) < lin:
            self.linhas.append([])
        row = self.linhas[lin - 1]
        while len(row) < col - 1 + len(values[0]):
            row.append("")
        row[col - 1:col - 1 + len(values[0])] = list(values[0])

    def batch_update(self, blocos, **_):
        for b in blocos:
            self.update(range_name=b["range"], values=b["values"])

    def append_row(self, values, **_):
        self.linhas.append(list(values))
        self.col_count = max(self.col_count, len(values))

    def add_cols(self, n):
        self.col_count += n


class Spreadsheet:
    def __init__(self):
        self.abas = {}

    def worksheet(self, nome):
        if nome not in self.abas:
            raise WorksheetNotFound(nome)
        return self.abas[nome]

    def add_worksheet(self, title, rows, cols):
        if title in self.abas:
            raise ValueError("já existe")
        self.abas[title] = Worksheet(title, cols)
        return self.abas[title]


PLANILHA = Spreadsheet()


class Client:
    def open_by_key(self, key):
        return PLANILHA


def authorize(creds):
    return Client()


# Pré-carga FICTÍCIA: dois clientes Aqua Gestão. Alfa tem visita em setembro de 2026
# (fora do mês padrão, para exercitar a importação automática com mês vazio); Beta tem
# visita numa terça-feira do mês corrente (importação automática normal).
_hoje = _dt.date.today()
_terca = _dt.date(_hoje.year, _hoje.month, 1)
while _terca.weekday() != 1:
    _terca += _dt.timedelta(days=1)
_cli = PLANILHA.add_worksheet("👥 Clientes", 100, 15)
_cli.linhas = [
    ["", "ID", "Nome", "Vol", "Tel", "Contato", "Endereço", "", "Status", "", "", "", "Empresa", "CNPJ", "CEP"],
    ["", "C1", "Condomínio Alfa Fictício", "50", "", "Síndico Alfa", "Rua A, 1", "", "Ativo", "", "", "", "Aqua Gestão", "", ""],
    ["", "C2", "Residencial Beta Fictício", "40", "", "Síndico Beta", "Rua B, 2", "", "Ativo", "", "", "", "Aqua Gestão", "", ""],
]
_vis = PLANILHA.add_worksheet("🔬 Visitas", 100, 23)
_vis.linhas = [
    ["", "ID", "Data", "", "Condomínio"] + [""] * 18,
    ["", "V1", "29/09/2026", "", "Condomínio Alfa Fictício", "7,4", "1,5", "1,8", "90", "", "0"]
    + [""] * 8 + ["Concluída", "Operador Fictício", "", ""],
    ["", "V2", _terca.strftime("%d/%m/%Y"), "", "Residencial Beta Fictício", "7,2", "2,0", "2,1", "100", "", "30"]
    + [""] * 8 + ["Concluída", "Operador Fictício", "", ""],
]

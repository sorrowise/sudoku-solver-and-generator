import xlwings as xw
from random import choice
from pickle import load, dump
from solver import solve_sudoku
from os.path import join,dirname


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets["solver"]
    def f(x): return int(x) if x != None else 0
    grid = [[f(x) for x in row] for row in sheet.range("E4:M12").value]
    ans = list(solve_sudoku((3, 3), grid))[0]
    sheet.range("U4:AC12").value = ans


@xw.sub
def solve4():
    wb = xw.Book.caller()
    sheet = wb.sheets["4-grid"]
    def f(x): return int(x) if x != None else 0
    grid = [[f(x) for x in row] for row in sheet.range("F4:I7").value]
    ans = list(solve_sudoku((2, 2), grid))[0]
    sheet.range("P4:S7").value = ans


@xw.sub
def solve6():
    wb = xw.Book.caller()
    sheet = wb.sheets["6-grid"]
    def f(x): return int(x) if x != None else 0
    grid = [[f(x) for x in row] for row in sheet.range("D3:I8").value]
    ans = list(solve_sudoku((2, 3), grid))[0]
    sheet.range("P3:U8").value = ans


def easy():
    wb = xw.Book.caller()
    sheet = wb.sheets["solver"]
    path = join(dirname(__file__),"puzzles","easy.pkl")
    with open(path, 'rb') as f:
        puzzles = load(f)
    puzzle = choice(puzzles)
    sheet.range("E4:M12").value = puzzle


def medium():
    wb = xw.Book.caller()
    sheet = wb.sheets["solver"]
    path = join(dirname(__file__),"puzzles","medium.pkl")
    with open(path, 'rb') as f:
        puzzles = load(f)
    puzzle = choice(puzzles)
    sheet.range("E4:M12").value = puzzle


def hard():
    wb = xw.Book.caller()
    sheet = wb.sheets["solver"]
    path = join(dirname(__file__),"puzzles","hard.pkl")
    with open(path, 'rb') as f:
        puzzles = load(f)
    puzzle = choice(puzzles)
    sheet.range("E4:M12").value = puzzle


def expert():
    wb = xw.Book.caller()
    sheet = wb.sheets["solver"]
    path = join(dirname(__file__),"puzzles","expert.pkl")
    with open(path, 'rb') as f:
        puzzles = load(f)
    puzzle = choice(puzzles)
    sheet.range("E4:M12").value = puzzle


def evil():
    wb = xw.Book.caller()
    sheet = wb.sheets["solver"]
    path = join(dirname(__file__),"puzzles","evil.pkl")
    with open(path, 'rb') as f:
        puzzles = load(f)
    puzzle = choice(puzzles)
    sheet.range("E4:M12").value = puzzle


@xw.sub
def savepuzzle(path):
    wb = xw.Book.caller()
    sheet = wb.sheets["solver"]
    def f(x): return int(x) if x != None else 0
    grid = [[f(x) for x in row] for row in sheet.range("E4:M12").value]
    with open(path+"pkl", "wb") as f:
        dump(grid, f)


@xw.sub
def loadpuzzle(path):
    wb = xw.Book.caller()
    sheet = wb.sheets["solver"]
    with open(path, "rb") as f:
        grid = load(f)
    sheet.range("E4:M12").value = grid


if __name__ == "__main__":
    xw.Book("sudoku.xlsm").set_mock_caller()
    main()

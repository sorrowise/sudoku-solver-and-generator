from random import choice, sample


def construct_puzzle_solution():
    while True:
        try:
            puzzle = [[0]*9 for i in range(9)]
            rows = [set(range(1, 10)) for i in range(9)]
            columns = [set(range(1, 10)) for i in range(9)]
            squares = [set(range(1, 10)) for i in range(9)]
            for i in range(9):
                for j in range(9):
                    choices = rows[i].intersection(
                        columns[j]).intersection(squares[(i//3)*3 + j//3])
                    res = choice(list(choices))

                    puzzle[i][j] = res

                    rows[i].discard(res)
                    columns[j].discard(res)
                    squares[(i//3)*3 + j//3].discard(res)
            return puzzle
        except IndexError:
            pass


def mask(square, left=17):
    idxs = sample(range(81), k=81-left)
    for idx in idxs:
        x, y = idx//9, idx % 9
        square[x][y] = 0
    return square

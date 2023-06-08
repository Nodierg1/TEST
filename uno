import math
import random
import docx
from docx.shared import Pt, Cm
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from tqdm import tqdm

def generate_sudoku(size=9, difficulty='medium'):
    def is_possible(grid, x, y, n):
        for i in range(size):
            if grid[x][i] == n or grid[i][y] == n:
                return False
        x0 = (x // box_size) * box_size
        y0 = (y // box_size) * box_size
        for i in range(box_size):
            for j in range(box_size):
                if grid[x0 + i][y0 + j] == n:
                    return False
        return True

    def solve(grid):
        for x in range(size):
            for y in range(size):
                if grid[x][y] == 0:
                    for n in random.sample(range(1, size + 1), k=size):
                        if is_possible(grid, x, y, n):
                            grid[x][y] = n
                            if solve(grid):
                                return True
                            grid[x][y] = 0
                    return False
        return True

    def generate_puzzle(grid):
        if difficulty == 'easy':
            attempts = 3
        elif difficulty == 'medium':
            attempts = 5
        else:
            attempts = 7

        while attempts > 0:
            row = random.randint(0, size - 1)
            col = random.randint(0, size - 1)
            while grid[row][col] == 0:
                row = random.randint(0, size - 1)
                col = random.randint(0, size - 1)
            backup = grid[row][col]
            grid[row][col] = 0

            copy_grid = []
            for r in range(size):
                copy_grid.append([])
                for c in range(size):
                    copy_grid[r].append(grid[r][c])

            counter = 0
            solve(copy_grid)
            if counter > 1:
                grid[row][col] = backup
                attempts -= 1

    box_size = int(math.sqrt(size))
    grid = [[0 for _ in range(size)] for _ in range(size)]
    solve(grid)
    generate_puzzle(grid)
    return grid

def create_sudoku_docx_pdf(num_pages=200, size=9, difficulty='medium', with_answers=False, output_format='word'):
    doc = docx.Document()
    doc.sections[0].page_height = Cm(23.5)
    doc.sections[0].page_width = Cm(19.05)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(16)

    pbar = tqdm(total=num_pages, desc='Generando sudokus', unit='sudoku')

    for i in range(num_pages):
        board = generate_sudoku(size=size, difficulty=difficulty)
        table = doc.add_table(rows=size, cols=size)
        table.style = 'Table Grid'
        table.autofit = False
        table_borders = table._element.xpath('.//w:tblBorders/*')
        for border in table_borders:
            border.set(docx.oxml.ns.qn('w:sz'), '8')
        cell_size = 1.5 * Cm(19.05) / size / Cm(1)
        for row in table.rows:
            row.height = Pt(cell_size)
        for col in table.columns:
            col.width = Pt(cell_size)
        for row in range(size):
            for col in range(size):
                cell = table.cell(row, col)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                p = cell.paragraphs[0]
                p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(str(board[row][col]) if random.random() < 0.5 else '')
                run.font.color.rgb = docx.shared.RGBColor(0, 0, 0)
                run.font.size = Pt(20)
        if with_answers:
            p = doc.add_paragraph()
            p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
            p.page_break_before = True
            run = p.add_run('Respuestas:')
            run.font.size = Pt(16)
            table = doc.add_table(rows=size, cols=size)
            table.style = 'Table Grid'
            table.autofit = False
            cell_size = 1.5 * Cm(19.05) / size / Cm(1) / 10
            for row in table.rows:
                row.height = Pt(cell_size)
            for col in table.columns:
                col.width = Pt(cell_size)
            for row in range(size):
                for col in range(size):
                    cell = table.cell(row, col)
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    p = cell.paragraphs[0]
                    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run(str(board[row][col]))
                    run.font.color.rgb = docx.shared.RGBColor(0, 0, 0)
                    run.font.size = Pt(8)
        else:
            if i != num_pages - 1:
                doc.add_page_break()
        pbar.update(1)

    pbar.close()

    if output_format == 'word':
        doc.save('sudokus.docx')
        print('Sudokus generados y guardados en el archivo "sudokus.docx".')
    elif output_format == 'pdf':
        from docx2pdf import convert
        doc.save('sudokus.docx')
        convert("sudokus.docx", "sudokus.pdf")
        print('Sudokus generados y guardados en los archivos "sudokus.docx" y "sudokus.pdf".')

size_options = ['6x6', '9x9', '16x16', '25x25', '36x36']
difficulty_options = ['easy', 'medium', 'hard']
output_format_options = ['word', 'pdf']

size_str = input(f'Seleccione el tamaño del sudoku ({", ".join(size_options)}): ')
while size_str not in size_options:
    size_str = input(f'Opción inválida. Seleccione el tamaño del sudoku ({", ".join(size_options)}): ')
size = int(size_str.split('x')[0])

num_pages_str = input('Ingrese la cantidad de sudokus a generar: ')
while not num_pages_str.isdigit():
    num_pages_str = input('Entrada inválida. Ingrese la cantidad de sudokus a generar: ')
num_pages = int(num_pages_str)

difficulty_str = input(f'Seleccione el nivel de dificultad ({", ".join(difficulty_options)}): ')
while difficulty_str not in difficulty_options:
    difficulty_str = input(f'Opción inválida. Seleccione el nivel de dificultad ({", ".join(difficulty_options)}): ')
difficulty = difficulty_str

with_answers_str = input('¿Incluir soluciones? (S/N): ')
while with_answers_str not in ['S', 'N']:
    with_answers_str = input('Entrada inválida. ¿Incluir soluciones? (S/N): ')
with_answers = True if with_answers_str == 'S' else False

output_format_str = input(f'Seleccione el formato de salida ({", ".join(output_format_options)}): ')
while output_format_str not in output_format_options:
    output_format_str = input(f'Opción inválida. Seleccione el formato de salida ({", ".join(output_format_options)}): ')
output_format = output_format_str

create_sudoku_docx_pdf(num_pages=num_pages, size=size, difficulty=difficulty, with_answers=with_answers, output_format=output_format)

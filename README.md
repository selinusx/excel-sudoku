a prototype sudoku project just for myself,extremely beginner level in excel.
# excel sudoku (vba-only)

## what is this
plain-text vba module (.bas) for solving a 9×9 sudoku in excel. no binary files, no personal metadata.

## how to use
1. open excel and create a new workbook.
2. make a 9×9 grid at **A1:I9** (optional: add borders).
3. press **alt+f11** (mac: fn+alt+f11) to open the vba editor.
4. **insert → module**, then paste the contents of `sudoku_solver.bas`.
5. return to excel, put your puzzle into **A1:I9** (leave blanks empty).
6. **developer → macros → SolveSudoku → run** (or assign to a button).

- `SolveSudoku` fills the grid.
- `ClearGrid` empties A1:I9.
- `LoadExample` writes a sample puzzle.

## notes
- this backtracking solver expects a valid puzzle with at least one solution.
- blanks must be empty cells (not zeros as text).
- you can commit only `.bas` (this file) to github for anonymity — no `.xlsm` needed.

## license
mit

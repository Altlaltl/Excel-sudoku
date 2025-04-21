# Excel-sudoku
Sudoku with excel VBA macro


Okay, imagine a 9x9 grid, our Sudoku board.

Creating the Solved Grid:

    Sub-squares First: Think of the grid as nine 3x3 little squares inside. We try to fill each of these small squares with the numbers 1 to 9, one number at a time.

    +-------+-------+-------+
    | 1 2 3 |       |       |  <- We try to put '1' in this first 3x3,
    | 4 5 6 |       |       |     then maybe '2', and so on.
    | 7 8 9 |       |       |
    +-------+-------+-------+
    |       |       |       |
    |       |       |       |
    |       |       |       |
    +-------+-------+-------+
    |       |       |       |
    |       |       |       |
    |       |       |       |
    +-------+-------+-------+

    Random Spot, Rule Check: For each number (say, we're trying to place '3'), we pick a random empty spot within the current 3x3 square. Before we put it there, we check:
        Same Row? Is '3' already in that horizontal line?
        Same Column? Is '3' already in that vertical line?

    +-------+-------+-------+
    | 1 2 . |       |       |  <- Trying to put '3' here (.), check row
    | 4 5 . |       |       |     and column.
    | 7 8 . |       |       |
    +-------+-------+-------+

    If It's Okay, Place It: If '3' isn't in the same row or column, we put it in the random spot.

    Stuck? Go Back: If we try many times to place a number in a 3x3 square and can't find a valid spot, it means we've hit a dead end. We then go back to a number we placed earlier and try a different spot for it. This "going back" is called backtracking. Sometimes we might even have to erase a few numbers to find a way forward.

    Repeat for All Numbers, All Squares: We keep doing this for numbers 1 through 9, and for all nine 3x3 squares. Eventually, we get a fully filled Sudoku grid that follows all the rules. This is our solved puzzle.

Creating the Puzzle (Removing Numbers):

    Start with the Solved Grid: Now we have our complete Sudoku.

    +-------+-------+-------+
    | 1 2 3 | 4 5 6 | 7 8 9 |
    | 4 5 6 | 7 8 9 | 1 2 3 |
    | 7 8 9 | 1 2 3 | 4 5 6 |
    +-------+-------+-------+
    | 2 3 4 | 5 6 7 | 8 9 1 |
    | 5 6 7 | 8 9 1 | 2 3 4 |
    | 8 9 1 | 2 3 4 | 5 6 7 |
    +-------+-------+-------+
    | 3 4 5 | 6 7 8 | 9 1 2 |
    | 6 7 8 | 9 1 2 | 3 4 5 |
    | 9 1 2 | 3 4 5 | 6 7 8 |
    +-------+-------+-------+

    Pick a Cell to Remove: We randomly choose a cell that has a number in it.

    +-------+-------+-------+
    | 1 2 3 | 4 5 6 | 7 8 9 |
    | 4 5 6 | 7 8 9 | 1 2 3 |
    | 7 8 9 | 1 2 3 | 4 5 6 |
    +-------+-------+-------+
    | 2 3 4 | 5 6 7 | 8 9 1 |
    | 5 6 7 | 8 9 1 | 2 3 4 |  <- Let's say we pick this '2'
    | 8 9 1 | 2   4 | 5 6 7 |
    +-------+-------+-------+
    | 3 4 5 | 6 7 8 | 9 1 2 |
    | 6 7 8 | 9 1 2 | 3 4 5 |
    | 9 1 2 | 3 4 5 | 6 7 8 |
    +-------+-------+-------+

    Check for Uniqueness: Now, we imagine this cell is empty. We then use a special "solver" (the HasUniqueSolution part of the code) to see if this new, incomplete puzzle still has only one possible correct answer.

        One Answer? Good: If there's still only one way to fill the empty spots correctly, we leave the cell empty. We've successfully made the puzzle a little harder.

        More Than One Answer? Oops: If filling the empty spots could lead to multiple different valid Sudoku solutions, it means our puzzle is now too easy or ambiguous. We put the original number back in that cell.

    Repeat Until Difficulty Reached: We keep picking random filled cells and trying to remove them, always checking if the puzzle still has only one solution. We stop when we've removed enough numbers to reach the difficulty level we want (e.g., remove more for "Difficult" than "Easy").

Solving Algorithm :


The Logic: Step-by-Step Deduction 

    Scan for Obvious Singletons: Look at each empty cell. For that cell, check:
        The Row: Are there any numbers already in that horizontal line?
        The Column: Are there any numbers already in that vertical line?
        The 3x3 Square: Are there any numbers already in that little 3x3 box?

    +-------+-------+-------+
    | 1 2 . |       |       |  <- Looking at '.',
    | 4 5 . |       |       |     '1', '2' are in the row.
    | 7 8 . |       |       |     '1', '4', '7' are in the column (if filled).
    +-------+-------+-------+

    Eliminate Possibilities: The numbers you see in the same row, column, and 3x3 square cannot be the number in the empty cell. For our example '.', if '1', '2', '4', '5', '7', '8' are already present in the related row, column, or 3x3 square, then the only possibilities left for '.' are '3', '6', and '9'.

    The "Only One Left" Rule: If, after eliminating the impossible numbers, there's only one number left that could possibly go in the empty cell, then that's the answer! You fill it in.

    +-------+-------+-------+
    | 1 2 ? |       |       |  <- If '4', '5', '6', '7', '8', '9' are in the
    | 4 5 . |       |       |     related row/column/square, then '?' MUST be '3'.
    | 7 8 . |       |       |
    +-------+-------+-------+

    Repeat the Scan: Keep scanning the grid, looking for these "only one left" situations in the empty cells. Each time you fill in a number, it might create new "only one left" opportunities in other empty cells.

    More Advanced Deductions (If Simple Scanning Doesn't Solve It): Sometimes, just looking at individual cells isn't enough. You need to look at groups of cells:

        Hidden Singles: In a row, column, or 3x3 square, a certain number might only be a possible candidate in one specific empty cell within that group, even if it's not the only possibility for that cell.

        Row: . 2 . | . . 5 | . . .
        Candidates for first '.': 1, 3, 6
        Candidates for third '.': 1, 4, 6
        If '1' can *only* be in the first '.' in this row, then the first '.' must be '1'.

        Pairs, Triples, Quads (Naked and Hidden): If two empty cells in a row, column, or 3x3 square have the exact same two possible candidate numbers, then those two numbers can be eliminated as possibilities from other empty cells in that same row, column, or square. Similar logic applies to triples (three cells, three same candidates) and quads. "Hidden" versions are where the pair/triple/quad of candidates are the only candidates appearing in that set of cells within the row/column/square.

    Guessing (If Absolutely Stuck - More Like the Computer's Backtracking): If you reach a point where you can't make any more logical deductions, you might have to make a guess.
        Pick an Empty Cell with Few Possibilities: Choose an empty cell that only has two or three possible numbers.
        Make a Tentative Guess: Try one of the possibilities and fill it in.
        Continue Solving: Keep trying to solve the puzzle based on your guess.
        Find a Contradiction? Wrong Guess: If your guess leads to a situation where you break a Sudoku rule (e.g., two of the same number in a row), then your guess was wrong. You backtrack, erase your guess, and try one of the other possibilities for that cell.

    This guessing step is very similar to the backtracking the computer uses when creating a puzzle. When solving, a good human solver tries to avoid guessing as much as possible and rely on logical deductions. However, for very difficult puzzles, guessing might become necessary.

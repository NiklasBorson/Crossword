# Crossword

This repo contains a PowerShell script for authoring crossword puzzles.

## Usage

To create a new crossword puzzle, one first invokes the script as follows:

```pwsh
.\Build-Crossword.ps1 -Title MyPuzzle
```

The script generates the following files. The file names below assume the
title is `MyPuzzle`.

| File name               | Description
|-------------------------|---------------------------------------------------------|
| `MyPuzzle-Grid.txt`     | _Grid file_ representing the contents of the crossword. |
| `MyPuzzle-Words.txt`    | _Word list file_ specifying words and their clues.      |
| `MyPuzzle-Puzzle.htm`   | Output HTML file for the crossword.                     |
| `MyPuzzle-Answers.htm`  | Output HTML file for the answer key.                    |

The script opens the grid file in Notepad so you can edit the grid -- that is,
specify what letter is in each cell.

The script then opens the word list file so you can edit the clues for each word.

Finally, the script generates HTML files representing the puzzle and answer key.

You can run the script more than once with the same title to refine your
crossword puzzle by editing the grid and/or clues from the previous run.

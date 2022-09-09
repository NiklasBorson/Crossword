# Crossword

This repo contains a PowerShell script for authoring crossword puzzles.

## Usage

To create a new crossword puzzle, one first invokes the script as follows:

```pwsh
.\Build-Crossword.ps1 -Title MyPuzzle
```

The script generates the following files. The file names below assume the
title is `MyPuzzle`.

| File name               | Description                                 |
|-------------------------|---------------------------------------------|
| `MyPuzzle-Grid.txt`     | Grid file representing the contents of your |
|                         | crossword, i.e., what letter in each cell.  |
|                         | The script opens this file in Notepad so    |
|                         | you can edit the grid.                      |
|-------------------------|---------------------------------------------|
| `MyPuzzle-Words.txt`    | Word list file specifying the ACROSS and    |
|                         | DOWN words in the crossword and the clue    |
|                         | for each. The script opens this file in     |
|                         | Notepad so you can edit the clues.          |
|-------------------------|---------------------------------------------|
| `MyPuzzle-Puzzle.htm`   | Output HTML file for the crossword.         |
|-------------------------|---------------------------------------------|
| `MyPuzzle-Answers.htm`  | Output HTML file for the answer key.        |

You can run the script more than once with the same title to refine your
crossword puzzle by editing the grid and/or clues.

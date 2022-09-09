param (
    [Parameter(Mandatory=$True)][string] $Title <# Crossword title and base file name #>,
    [Parameter(Mandatory=$False)][switch] $Unattended <# Run without user interaction #>
)

<# Gets the contents of a text file with comments, blank lines, and trailing spaces removed. #>
function Get-Text([string] $FileName) {
    Get-Content $FileName | Where-Object { $_ -and ($_[0] -ne '#') } | ForEach-Object { $_.TrimEnd() }
}

<# Returns the CSS markup for generated HTML files. #>
function Get-Css {
    'h1 { font-family:Calibri; font-size:24pt; font-weight:bold }'
    'table.puzzle, table.puzzle td { border:1px solid black; border-collapse:collapse; }'
    'table.puzzle td { width:26pt; height:28pt; margin:0; padding: 0; vertical-align:top; }'
    'td.empty { background-color:black; print-color-adjust:exact; -webkit-print-color-adjust:exact; }'
    'div.number { font-size:8pt; font-weight:bold; text-align:left; position:relative; left:1pt; top:0; height:0; }'
    'div.letter { font-size:20pt; font-family:Consolas, monospace; text-align:center; margin-top:3pt; }'
    'table.clues { margin-top:12pt; }'
    'table.clues th, table.clues td { width:3.5in; text-align:left; vertical-align:top; }'
    'table.clues th { font-family:Calibri; font-size:16pt; font-weight:bold; }'
    'table.clues td { font-family:Calibri; font-size:12pt; font-weight:normal; }'
    'p { margin-top:0; margin-bottom:4pt; }'
}

<# Contains information about a numbered cell in the crossword #>
class Cell {
    [int] $X <# zero-based column index #>
    [int] $Y <# zero-based row index #>
    [int] $CellNumber <# unique cell number #>
    [string] $AcrossWord <# across word beginning at this cell, if any #>
    [string] $DownWord <# down word beginning at this cell, if any #>
    [string] $AcrossClue <# clue associated with across word, if any #>
    [string] $DownClue <# clue associated with down word, if any #>
}

<# Mapping of words and cells to clues #>
class ClueMap {

    # CellMap stores clues indexed by cell number, direction, AND word. This is so the
    # same word appearing at different places in the puzzle can be differentiated. E.g.,
    # "ANSWER" at 1-ACROSS can have a different clue than "ANSWER" at 5-DOWN.
    $CellMap = @{}

    # WordMap stores clues indexed by word only. This is used as a backup to CellMap
    # so clues can still be retrieved if the grid is modified and cells renumbered.
    $WordMap = @{}

    <# Add a clue to the clue map #>
    [void] Add([int] $cellNumber, [string] $direction, [string] $word, [string] $clue) {
        $this.CellMap["$cellNumber$direction-$word"] = $clue
        $this.WordMap[$word] = $clue
    }

    <# Find a clue in the clue map #>
    [string] Find([int] $cellNumber, [string] $direction, [string] $word) {
        $clue = $this.CellMap["$cellNumber$direction-$word"]
        if (-not $clue) {
            $clue = $this.WordMap[$word]
        }
        return $clue
    }
}

<# Represents the contents of a crossword puzzle #>
class Grid {
    [string[]] $Rows
    [int] $RowCount
    [int] $ColCount
    [Cell[]] $NumberedCells

    <# Returns the letter in a cell or $Null for a blank cell #>
    [string] GetCell([int] $x, [int] $y) {
        if (($x -ge 0) -and ($y -ge 0)) {
            if ($y -lt $this.Rows.Length) {
                $row = $this.Rows[$y]
                if ($x -lt $row.Length) {
                    $cell = $row[$x]
                    if (($cell -ge 'A') -and ($cell -le 'Z')) {
                        return $cell
                    }
                }
            }
        }
        return $Null
    }

    <# Returns the ACROSS word starting at the specified cell position, or $Null if none. #>
    [string] GetAcrossWord([int] $x, [int] $y) {
        $word = ''

        if (-not $this.GetCell($x - 1, $y)) {
            for ($i = $x; $i -lt $this.ColCount; $i++) {
                $ch = $this.GetCell($i, $y)
                if ($ch) {
                   $word += $ch
                }
                else {
                    break
                }
            }
        }

        if ($word.Length -lt 2) {
            $word = $Null
        }

        return $word
    }

    <# Returns the DOWN word starting at the specified cell position, or $Null if none. #>
    [string] GetDownWord([int] $x, [int] $y) {
        $word = ''

        if (-not $this.GetCell($x, $y - 1)) {
            for ($i = $y; $i -lt $this.RowCount; $i++) {
                $ch = $this.GetCell($x, $i)
                if ($ch) {
                    $word += $ch
                }
                else {
                    break
                }
            }
        }
        if ($word.Length -lt 2) {
            $word = $Null
        }

        return $word
    }

    <# Reads clues from a word list file #>
    [void] ReadClues([string] $WordListFile) {
        $direction = ''
        $clueMap = New-Object -TypeName ClueMap

        foreach ($line in (Get-Text $WordListFile)) {
            if ($line -eq 'ACROSS') {
                $direction = 'A'
            }
            elseif ($line -eq 'DOWN') {
                $direction = 'D'
            }
            elseif ($line -match ' *([0-9]+)\. ([A-Z]+): (.+)') {
                $cellNumber = [int]($matches[1])
                $word = $matches[2]
                $clue = $matches[3]
                $clueMap.Add($cellNumber, $direction, $word, $clue)
            }
        }

        foreach ($cell in $this.NumberedCells) {
            if ($cell.AcrossWord) {
                $clue = $clueMap.Find($cell.CellNumber, 'A', $cell.AcrossWord)
                if ($clue) {
                    $cell.AcrossClue = $clue
                }
            }
            if ($cell.DownWord) {
                $clue = $clueMap.Find($cell.CellNumber, 'D', $cell.DownWord)
                if ($clue) {
                    $cell.DownClue = $clue
                }
            }
        }
    }

    <# Generates an HTML file for this crossword #>
    [void] WriteHtml([string] $outputFileName, [string] $heading, [bool] $showLetters, [bool] $showClues) {

        # The file name passed to XmlWriter.Create must be a full path.
        $outputFileName = (New-Item -ItemType File -Force $outputFileName).FullName

        # Create the XmlWriter object with formatting
        $XmlSettings = New-Object System.Xml.XmlWriterSettings
        $XmlSettings.Indent = $True
        $XmlSettings.IndentChars = '  '
        $w = [System.Xml.XmlWriter]::Create($outputFileName, $XmlSettings)

        # Begin the document and top-level html element.
        $w.WriteStartDocument()
        $w.WriteStartElement('html', 'http://www.w3.org/1999/xhtml')
        $w.WriteAttributeString('dir', 'ltr')
        $w.WriteAttributeString('lang', 'en')

        # Write the head section
        $w.WriteStartElement('head')
        $w.WriteStartElement('title')
        $w.WriteString($heading)
        $w.WriteEndElement() # /title
        $w.WriteStartElement('style')
        $css = Get-Css
        $w.WriteString($css)
        $w.WriteEndElement() # /style
        $w.WriteEndElement() # /head

        $w.WriteStartElement('body') # /body
        $w.WriteStartElement('h1')
        $w.WriteString($heading)
        $w.WriteEndElement() # /h1

        # Determine the index and x,y coordinates of the next numbered cell.
        $cellIndex = 0
        $nextX = -1
        $nextY = -1
        if ($cellIndex -lt $this.NumberedCells.Length) {
            $nextX = $this.NumberedCells[$cellIndex].X
            $nextY = $this.NumberedCells[$cellIndex].Y
        }

        # Generate a table for the crossword grid.
        $w.WriteStartElement('table')
        $w.WriteAttributeString('class', 'puzzle')
        for ($y = 0; $y -lt $this.RowCount; $y++) {
            $w.WriteStartElement('tr')
            for ($x = 0; $x -lt $this.ColCount; $x++) {
                $w.WriteStartElement('td')

                $letter = $this.GetCell($x, $y)
                if ($letter) {
                    if (($x -eq $nextX) -and ($y -eq $nextY)) {
                        $number = [string]($this.NumberedCells[$cellIndex].CellNumber)
                        $cellIndex++
                        if ($cellIndex -lt $this.NumberedCells.Length) {
                            $nextX = $this.NumberedCells[$cellIndex].X
                            $nextY = $this.NumberedCells[$cellIndex].Y
                        }

                        $w.WriteStartElement('div')
                        $w.WriteAttributeString('class', 'number')
                        $w.WriteString($number)
                        $w.WriteEndElement() # /div
                    }

                    if ($showLetters) {
                        $w.WriteStartElement('div')
                        $w.WriteAttributeString('class', 'letter')
                        $w.WriteString($letter)
                        $w.WriteEndElement() # /div
                    }
                }
                else {
                    # Empty cell
                    $w.WriteAttributeString('class', 'empty')
                }

                $w.WriteEndElement() # /td
            }
            $w.WriteEndElement() # /tr
        }
        $w.WriteEndElement() # /table

        if ($showClues) {
            $w.WriteStartElement('table')
            $w.WriteAttributeString('class', 'clues')

            # Write the header row.
            $w.WriteStartElement('tr')
            $w.WriteStartElement('th')
            $w.WriteString('Across')
            $w.WriteEndElement() # /th
            $w.WriteStartElement('th')
            $w.WriteString('Down')
            $w.WriteEndElement() # /th
            $w.WriteEndElement() # /tr

            # Write the body row.
            $w.WriteStartElement('tr')
            $w.WriteStartElement('td')
            $this.NumberedCells | ForEach-Object {
                if ($_.AcrossClue) {
                    $w.WriteStartElement('p')
                    $w.WriteString("$($_.CellNumber). $($_.AcrossClue)")
                    $w.WriteEndElement()
                }
            }
            $w.WriteEndElement() # /td
            $w.WriteStartElement('td')
            $this.NumberedCells | ForEach-Object {
                if ($_.DownClue) {
                    $w.WriteStartElement('p')
                    $w.WriteString("$($_.CellNumber). $($_.DownClue)")
                    $w.WriteEndElement()
                }
            }
            $w.WriteEndElement() # /td
            $w.WriteEndElement() # /tr

            $w.WriteEndElement() # /table
        }

        $w.WriteEndElement() # /body

        # End the top-level html element and document.
        $w.WriteEndElement() # /html
        $w.WriteEndDocument()
        $w.Flush()
        $w.Close()
    }
}

<# Creates a new Grid object from a text file containing words arranged in a grid #>
function Get-Grid([string] $fileName) {
    $grid = New-Object -TypeName Grid
    $grid.Rows = Get-Text $fileName
    $grid.RowCount = $grid.Rows.Length
    $grid.ColCount = ($grid.Rows | Measure-Object -Maximum { $_.Length }).Maximum

    for ($y = 0; $y -lt $grid.RowCount; $y++) {
        for ($x = 0; $x -lt $grid.ColCount; $x++) {
            $across = $grid.GetAcrossWord($x, $y)
            $down = $grid.GetDownWord($x, $y)
            if ($across -or $down) {
                $cell = New-Object -TypeName Cell
                $cell.X = $x
                $cell.Y = $y
                $cell.CellNumber = $grid.NumberedCells.Length + 1
                $cell.AcrossWord = $across
                $cell.DownWord = $down
                $grid.NumberedCells += $cell
            }
        }
    }

    $grid
}

<# Gets the contents of a placeholder grid file #>
function Get-ExampleGrid {
    '# Create the grid for your crossword below using only capital'
    '# letters (A-Z) and spaces. Each line and column corresponds'
    '# to one cell of the final crossword puzzle, with spaces'
    '# representing blank cells.'
    '#'
    '# Lines beginning with "#" are comments and do not affect the'
    '# contents of the crossword puzzle.'
    '#'
    '# In the example below, 1-ACROSS is CROSSWORD, 2-DOWN is ROAD,'
    '# and so on. Cell numbers are not specified here because they'
    '# are assigned automatically later.'
    '#'
    'CROSSWORD'
    ' O   ARE'
    ' A  AGENT'
    ' D     T'
}

<# Gets the contents of a word list file from a Grid object #>
function Get-WordList([Grid] $grid) {
    '# Write the clues for your crossword by replacing the WRITE_CLUE_HERE'
    '# placeholder strings.'
    ''
    'ACROSS'

    foreach ($cell in $grid.NumberedCells) {
        if ($cell.AcrossWord) {
            $clue = $cell.AcrossClue ? $cell.AcrossClue : 'WRITE_CLUE_HERE'
            Write-Output " $($cell.CellNumber). $($cell.AcrossWord): $clue"
        }
    }

    ''
    'DOWN'

    foreach ($cell in $grid.NumberedCells) {
        if ($cell.DownWord) {
            $clue = $cell.DownClue ? $cell.DownClue : 'WRITE_CLUE_HERE'
            Write-Output " $($cell.CellNumber). $($cell.DownWord): $clue"
        }
    }
}

function Main {
    $GridFile = "$Title-Grid.txt"
    $WordListFile = "$Title-Words.txt"
    $PuzzleFile = "$Title-Puzzle.htm"
    $AnswerFile = "$Title-Answers.htm"

    # Generate the placeholder grid file if it doesn't already exist.
    if (-not (Test-Path $GridFile)) {
        Get-ExampleGrid | Set-Content $GridFile
    }

    # Prompt the user to edit grid, then read it.
    if (-not $Unattended) {
        Start-Process -Wait Notepad $GridFile
    }
    $grid = Get-Grid $GridFile

    # Read clues from the existing word list file if present.
    if (Test-Path $WordListFile) {
        $grid.ReadClues($WordListFile)
    }

    # Generate a word list and prompt the user to edit it.
    Get-WordList $grid | Set-Content $WordListFile
    if (-not $Unattended) {
        Start-Process -Wait Notepad $WordListFile
        $grid.ReadClues($WordListFile)
    }

    # Generate the HTML output files.
    $grid.WriteHtml($PuzzleFile, $Title, $False, $True)
    $grid.WriteHtml($AnswerFile, "$Title Answers", $True, $False)

    # Open the output files.
    if (-not $Unattended) {
        Start-Process $PuzzleFile
        Start-Process $AnswerFile
    }
}

Main

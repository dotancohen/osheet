
# Osheet

CLI application for browsing spreadsheets

File formats currently supported:
- XLS


## Known bugs

- Currently sizes display for the screen I use to develop on
- When passing the cursor outside the boundarides, crash
- Overflowing text highlights the trailing coloum separator
- Search crashes on wrap around to first search result


## Features to add

- Space on overflowing cell to show entire cell contents
- Scroll
- Resize columns
- Support CSV, TSV, and LibreOffice document formats


## Keys

### Navigation

`h` Left
`j` Down
`k` Up
`l` Right

### Operations

`/` Search
`n` Next search result
`<spacebar>` Expand cell to view full text

### Quit

`q` Quit


## Other applications to view spreadsheets on CLI

- [SC-IM](https://github.com/andmarti1424/sc-im): I would actually be using this if it supported arbitrary text search.


## Resources used to learn

### Python spreadsheet libraries

https://realpython.com/openpyxl-excel-spreadsheets-python/

### Curses

https://docs.python.org/3/howto/curses.html
https://gist.github.com/claymcleod/b670285f334acd56ad1c
https://docs.python.org/3/library/curses.html
https://www.devdungeon.com/content/curses-programming-python



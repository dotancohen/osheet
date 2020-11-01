#!/usr/bin/env python3

import openpyxl
import curses
import time

from pprint import pprint

file = "Example.xlsx"

cellsize = 12
col_sep = " | "
col_sep_size = len(col_sep)



def format_row(val_orig, size):
	val = str(val_orig) + " "*size
	if size < len(str(val_orig)):
		overflow = True
	else:
		overflow = False

	return (val[0:size], overflow, val_orig)



def get_yx(cell_y, cell_x):

	x = 2 + (cell_x-1) * (cellsize + col_sep_size)

	return (cell_y,x)



def main(stdscr):

	stdscr.border(0)

	workbook = openpyxl.load_workbook(filename=file)
	sheet = workbook.active

	row_num = 0
	cell_fulltext = []

	for row in sheet.iter_rows(min_row=0, max_row=50, min_col=0, max_col=10, values_only=True):
		row_num += 1
		col_num = 0
		row_fulltext = []
		for cell in row:
			if cell==None:
				cell=" "
			col_effective = (col_num*(cellsize+col_sep_size)) + 2
			(text,overflow,text_full) = format_row(cell, cellsize)
			if overflow:
				cell_formatting = curses.A_UNDERLINE
			else:
				cell_formatting = curses.A_NORMAL
			text += col_sep
			stdscr.addstr(row_num, col_effective, text, cell_formatting)
			row_fulltext.append(text_full)
			col_num += 1
		cell_fulltext.append(row_fulltext)

	stdscr.refresh()
	cell_y = 1
	cell_x = 1
	(y,x) = get_yx(cell_y, cell_x)
	stdscr.move(y, x)

	while True:
		ch = stdscr.getch()

		# Navigaton

		if ch == ord('h'):
			cell_x -= 1
			(y,x) = get_yx(cell_y, cell_x)
			stdscr.move(y, x)
		if ch == ord('j'):
			cell_y += 1
			(y,x) = get_yx(cell_y, cell_x)
			stdscr.move(y, x)
		if ch == ord('k'):
			cell_y -= 1
			(y,x) = get_yx(cell_y, cell_x)
			stdscr.move(y, x)
		if ch == ord('l'):
			cell_x += 1
			(y,x) = get_yx(cell_y, cell_x)
			stdscr.move(y, x)

		# Operatons

		if ch == ord(' '):
			win = curses.newwin(5, cellsize*2, y, x)
			win.border(0)
			win.addstr(1, 1, cell_fulltext[cell_y-1][cell_x-1], curses.A_NORMAL)
			win.refresh()
			win.getch()
			del win
			stdscr.touchwin()
			stdscr.refresh()

		# Quit

		if ch == ord('q'):
			break

	return



if __name__=="__main__":
	curses.wrapper(main)

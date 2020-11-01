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



def main(stdscr):

	stdscr.border(0)

	workbook = openpyxl.load_workbook(filename=file)
	sheet = workbook.active

	row_num = 1

	for row in sheet.iter_rows(min_row=0, max_row=50, min_col=0, max_col=10, values_only=True):
		row_num += 1
		col_num = 0
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
			col_num += 1

	stdscr.refresh()
	cursor_y = 2
	cursor_x = 2
	stdscr.move(cursor_y, cursor_x)

	while True:
		ch = stdscr.getch()

		if ch == ord('h'):
			cursor_x -= cellsize + col_sep_size
			stdscr.move(cursor_y, cursor_x)
		if ch == ord('j'):
			cursor_y += 1
			stdscr.move(cursor_y, cursor_x)
		if ch == ord('k'):
			cursor_y -=1
			stdscr.move(cursor_y, cursor_x)
		if ch == ord('l'):
			cursor_x += cellsize + col_sep_size
			stdscr.move(cursor_y, cursor_x)
		if ch == ord('q'):
			break

	return



if __name__=="__main__":
	curses.wrapper(main)

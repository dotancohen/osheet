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
	return val[0:size]



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
			text = format_row(cell, cellsize)
			text += col_sep
			stdscr.addstr(row_num, col_effective, text, curses.A_NORMAL)
			col_num += 1

	stdscr.refresh()

	ch = stdscr.getch()


	return



if __name__=="__main__":
	curses.wrapper(main)

#!/usr/bin/env python3

import argparse
import sys
import openpyxl
import curses
import time
import re


version = 0.1

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


def search(scr, cell_fulltext):

	matches = [] # y,x

	search_str = scr.getstr().decode("ascii")


	for cell_y,row in enumerate(cell_fulltext):
		for cell_x,cell in enumerate(row):
			if not type(cell)==type("string"):
				continue
			if re.search(search_str, cell, re.I):
				matches.append( (cell_y+1,cell_x+1) )

	return matches


def match_next(matches, match_index, scr):
	(cell_y, cell_x) = matches[match_index]
	(y,x) = get_yx(cell_y, cell_x)
	scr.move(y, x)
	return (cell_y, cell_x)


def main(stdscr, file):

	stdscr.border(0)
	matches = []
	match_index = 0

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

		if ch == ord('/'):
			matches = search(stdscr, cell_fulltext)
			match_index = 0;
			if 0 < len(matches):
				(cell_y, cell_x) = match_next(matches, match_index, stdscr)
				match_index += 1;
		if ch == ord('n'):
			if 0 < len(matches):
				(cell_y, cell_x) = match_next(matches, match_index, stdscr)
				match_index += 1;

		# Quit

		if ch == ord('q'):
			break

	return



if __name__=="__main__":
	parser = argparse.ArgumentParser(description="CLI application for browsing spreadsheets")
	parser.add_argument('file', type=str, help="Spreadsheet file to browse")
	args = parser.parse_args()

	curses.wrapper(main, args.file)

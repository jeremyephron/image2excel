#!/usr/bin/env python3
"""
File: image2excel.py
--------------------
This script converts an image into an Excel spreadsheet.

Usage: ./image2excel.py image [-o OUTPUT]

where image is the path to your image and output is your desired output file
path. If output is not specified, the file will take the name of your image 
and be placed in the current directory.

Look at the image at 20% zoom for the best quality. Your image will be scaled 
down to fit maximally inside the MAX_SIZE specified below for the best 
viewing experience.

Each pixel corresponds to three vertically stacked cells, whose fill colors
are conditionally formatted to be black if the value is 0, and red (0x0000FF),
green (0x00FF00), or blue (0xFF0000), respectively, if the value is 255.

Acknowledgments: Thanks to Matt Parker for the image-to-spreadsheet idea
(http://standupmaths.com/).

"""

import argparse
from pathlib import Path
import string

import cv2
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font

MAX_SIZE = (200, 477)  # max image dimension (height, width)
                       # these dims make the image look great at 20% zoom
ROW_HEIGHT = 5
COL_WIDTH = 2.45  # sets column width to 1.67, which is optimal  (not sure why)
FONT_SIZE = 2


def parse_args():
    parser = argparse.ArgumentParser(description='Convert an image to an '
                                                 'excel spreadsheet.')
    parser.add_argument('image', type=str, help='the filepath of the image')
    parser.add_argument('-o', '--output', required=False, type=str, 
                        help='the xlsx filepath to output')
    return vars(parser.parse_args())


def main(image, output=None):
    # Load the image
    img = cv2.imread(image)

    if img is None:
        raise Exception(f'Failed to load image "{image}".')

    # Resize the image so you can see it will fit on the screen
    img = bound_image_to_size(img, MAX_SIZE)

    # Create a new workbook
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title='Image')
    
    # Write the image data to the worksheet
    write_img_to_ws(img, ws)

    # Apply conditional formatting for the RGB colors
    apply_colors(ws, img.shape[0] * img.shape[2], img.shape[1])

    # Name the output file after image if not specified
    if output is None:
        output = Path(image).stem + '.xlsx'

    # Save the image
    wb.save(output)


def bound_image_to_size(img, size):
    """
    Bounds an image by the specified size, scaling the image down until it is
    less than size.

    Args:
        img: the image to scale down.
        size: the (height, width) bounds.

    """

    # Image fits
    if img.shape[0] <= size[0] and img.shape[1] <= size[1]:
        return img

    # Height is limiting
    if (size[0] / max(img.shape[0], size[0]) 
            < size[1] / max(img.shape[1], size[1])):
        scale = size[0] / img.shape[0]

        # resize takes (width, height)
        return cv2.resize(img, (int(img.shape[1] * scale), size[0]))

    # Width is limiting
    else:
        scale = size[1] / img.shape[1]
        return cv2.resize(img, (size[1], int(img.shape[0] * scale)))


def write_img_to_ws(img, ws):
    """
    Writes an image array to the given worksheet.

    Args:
        img: the image array.
        ws: the write only worksheet.

    """
    
    # Set all row heights
    for i in range(1, img.shape[0] * img.shape[2] + 1):
        ws.row_dimensions[i].height = ROW_HEIGHT

    # Set all column widths
    for col in iter_col_names(img.shape[1]):
        ws.column_dimensions[col].width = COL_WIDTH

    # Helper function to create a cell with desired formatting
    def cell(val):
        cell = WriteOnlyCell(ws, value=val)
        cell.font = Font(name='Calibri', size=FONT_SIZE)
        cell.alignment = Alignment(vertical='center')
        return cell

    # Add all numbers to the worksheet
    for r in range(img.shape[0]):  # rows
        for channel in range(img.shape[2]):  # rgb channels
            ws.append([cell(x) for x in img[r, :, channel]])


def apply_colors(ws, n_rows, n_cols):
    """
    Apply colors to the worksheet by conditional formatting.

    Args:
        ws: the write only worksheet.
        n_rows: the number of rows in the worksheet.
        n_cols: the number of columns in the worksheet.

    """

    rgb = ['0000FF', '00FF00', 'FF0000']  # red, green, and blue in hex

    end_col = get_col_name(n_cols)
    for r in range(1, n_rows + 1):
        rule = ColorScaleRule(
            start_type='num', start_value=0, start_color='000000',
            end_type='num', end_value=255, end_color=rgb[(r - 1) % 3]
        )
        ws.conditional_formatting.add(f'A{r}:{end_col}{r}', rule)


def get_col_name(n):
    """
    Gets the column base alphabet name from number.

    Args:
        n: the column number (1-indexed)

    Returns:
        the column name in base alphabet ('AAZ').

    """

    col = ''
    while n > 0:
        col = string.ascii_uppercase[(n - 1) % 26] + col
        n = (n - 1) // 26

    return col


def iter_col_names(n_cols):
    """
    Iterates through column names.

    Args:
        n_cols: the number of columns to iterate through.

    Returns:
        the column name in base alphabet ('AAZ').

    """

    for i in range(1, n_cols + 1):
        yield get_col_name(i)


if __name__ == '__main__':
    main(**parse_args())

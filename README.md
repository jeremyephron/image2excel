Image2Excel
===========

Image2Excel is a Python script that converts your image into an Excel 
spreadsheet (an idea due to Matt Parker, look at
[Acknowledgments](#acknowledgements)).

Installation
------------

- Requires at least version Python3.6.

- Download the `image2excel.py` file and `requirements.txt`. You can also
  clone this repository (but you don't need any other files).

- Install requirements: `python3 -m pip install -r requirements.txt`. This 
  script requires `opencv-python` and `openpyxl`.

Usage
-----

Run `./image2excel.py image [-o OUTPUT]` where `image` is the path to your 
image and `OUTPUT` is your desired output file path (a xlsx file). If output 
is unspecified, the output file will take the name of your image.

After running, open up the output file and zoom out to 20% for the best 
viewing experience.

Screenshots
-----------

An image of Matt Parker:

![Image of Matt Parker](/screenshots/matt_parker.png)

Zoom 400%:

![Spreadsheet Zoom 400%](/screenshots/zoom_400.png)

Zoom 200%:

![Spreadsheet Zoom 200%](/screenshots/zoom_200.png)

Zoom 100%:

![Spreadsheet Zoom 100%](/screenshots/zoom_100.png)

Zoom 20%:

![Spreadsheet Zoom 20%](/screenshots/zoom_20.png)

Spreadsheet inception:

![Spreadsheet of Spreadsheet](/screenshots/spreadsheet_inception.png)

Notes
-----

- I found that the largest image that fits well in the spreadsheet on a 
  standard 13-inch laptop screen is 477x200. Input images will be scaled down
  to the largest size that is less than 477x200, preserving aspect-ratio.
  You can change this behavior by modifying the `MAX_SIZE` constant.

- The font size is set to 2, which is small enough to make the image look 
  great, but big enough to be read fully zoomed in. You can change the font 
  size by modifying the `FONT_SIZE` constant.

Acknowledgements
----------------

Thanks to [Matt Parker](http://standupmaths.com/) for great content and the 
spreadsheet idea. He has an [online converter](http://www.think-maths.co.uk/spreadsheet),
but this script produces a higher quality spreadsheet and supports larger 
image sizes.

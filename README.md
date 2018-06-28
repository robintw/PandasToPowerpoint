# PandasToPowerpoint (pd2ppt)
Python utility to take a Pandas DataFrame and create a Powerpoint table

An example:

```python
from pd2ppt import df_to_powerpoint
import pandas as pd

df = pd.DataFrame(
    {'District':['Hampshire', 'Dorset', 'Wiltshire', 'Worcestershire'],
     'Population':[25000, 500000, 735298, 12653],
     'Ratio':[1.56, 7.34, 3.67, 8.23]})

df_to_powerpoint(
    r"C:\Code\Powerpoint\test58.pptx", df, col_formatters=['', ',', '.2'],
    rounding=['', 3, ''])
```


## Installation

```bash
git clone https://github.com/robintw/PandasToPowerpoint.git
cd PandasToPowerpoint
pip install --upgrade pip # optional (depends on setup)
pip install -r requirements.txt
python setup.py install
```


## Documentation

### `df_to_table`
```
Converts a Pandas DataFrame to a PowerPoint table on the given
Slide of a PowerPoint presentation.

The table is a standard Powerpoint table, and can easily be modified with
the Powerpoint tools, for example: resizing columns, changing formatting etc.

Parameters
----------
slide: ``pptx.slide.Slide``
    slide object from the python-pptx library containing the slide on which
    you want the table to appear

df: pandas ``DataFrame``
   DataFrame with the data

left: int, optional
   Position of the left-side of the table, either as an integer in cm, or
   as an instance of a pptx.util Length class (pptx.util.Inches for
   example). Defaults to 4cm.

top: int, optional
   Position of the top of the table, takes parameters as above.

width: int, optional
   Width of the table, takes parameters as above.

height: int, optional
   Height of the table, takes parameters as above.

col_formatters: list, optional
   A n_columns element long list containing format specifications for each
   column. For example ['', ',', '.2'] does no special formatting for the
   first column, uses commas as thousands separators in the second column,
   and formats the third column as a float with 2 decimal places.

rounding: list, optional
   A n_columns element long list containing a number for each integer
   column that requires rounding that is then multiplied by -1 and passed
   to round(). The practical upshot of this is that you can give something
   like ['', 3, ''], which does nothing for the 1st and 3rd columns (as
   they aren't integer values), but for the 2nd column, rounds away the 3
   right-hand digits (eg. taking 25437 to 25000).

name: str, optional
   A name to be given to the table in the Powerpoint file. This is not
   displayed, but can help extract the table later to make further changes.

Returns
-------
pptx.shapes.graphfrm.GraphicFrame
    The python-pptx table (GraphicFrame) object that was created (which can
    then be used to do further manipulation if desired)
```

### `df_to_powerpoint`
```
Converts a Pandas DataFrame to a table in a new, blank PowerPoint
presentation.

Creates a new PowerPoint presentation with the given filename, with a single
slide containing a single table with the Pandas DataFrame data in it.

The table is a standard Powerpoint table, and can easily be modified with
the Powerpoint tools, for example: resizing columns, changing formatting
etc.

Parameters
----------
filename: Filename to save the PowerPoint presentation as

df: pandas ``DataFrame``
    DataFrame with the data

**kwargs:
    All other arguments that can be taken by ``df_to_table()`` (such as
    ``col_formatters`` or ``rounding``) can also be passed here.

Returns
-------
pptx.shapes.graphfrm.GraphicFrame
    The python-pptx table (GraphicFrame) object that was created (which can
    then be used to do further manipulation if desired)
```

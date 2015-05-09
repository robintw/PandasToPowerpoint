# PandasToPowerpoint
Python utility to take a Pandas DataFrame and create a Powerpoint table

An example:

```python
from PandasToPowerpoint import df_to_powerpoint

import pandas as pd

df = pd.DataFrame({'District':['Hampshire', 'Dorset', 'Wiltshire', 'Worcestershire'],
				   'Population':[25000, 500000, 735298, 12653],
				   'Ratio':[1.56, 7.34, 3.67, 8.23]})

df_to_powerpoint(r"C:\Code\Powerpoint\test58.pptx", df,
				  col_formatters=['', ',', '.2'], rounding=['', 3, ''])
```

## Documentation

### df_to_table
Converts a Pandas DataFrame to a PowerPoint table on the given
Slide of a PowerPoint presentation.

The table is a standard Powerpoint table, and can easily be modified with the Powerpoint tools,
for example: resizing columns, changing formatting etc.

Arguments:
 - slide: slide object from the python-pptx library containing the slide on which you want the table to appear
 - df: Pandas DataFrame with the data
 
Optional arguments:
 - col_formatters: A n_columns element long list containing format specifications for each column.
 For example `['', ',', '.2']` does no special formatting for the first column, uses commas as thousands separators
 in the second column, and formats the third column as a float with 2 decimal places.
 - rounding: A n_columns element long list containing a number for each integer column that requires rounding
 that is then multiplied by -1 and passed to `round()`. The practical upshot of this is that you can give something like
 `['', 3, '']`, which does nothing for the 1st and 3rd columns (as they aren't integer values), but for the 2nd column,
 rounds away the 3 right-hand digits (eg. taking 25437 to 25000).

### df_to_powerpoint
Converts a Pandas DataFrame to a table in a new, blank PowerPoint presentation.
    
Creates a new PowerPoint presentation with the given filename, with a single slide
containing a single table with the Pandas DataFrame data in it.

The table is a standard Powerpoint table, and can easily be modified with the Powerpoint tools,
for example: resizing columns, changing formatting etc.

Arguments:
 - filename: Filename to save the PowerPoint presentation as
 - df: Pandas DataFrame with the data

All other arguments that can be taken by `df_to_table()` (such as col_formatters or rounding) can also
be passed here.

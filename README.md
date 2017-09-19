# xlsxpandas: simple excel report drawing framework based on xlsxwriter and pandas

This is a simple framework for drawing excel reports. The main idea here is to have a plotter-like object (called a drawer),
that moves through an excel sheet matrix and draws objects that are supplied to it.

Additionaly `xlsxpandas` maps the two most important `pandas` data structures,
that is `Series` and `DataFrame`s, to corresponding drawable `xlsxpandas` elements,
so it is very easy to draw arbitrary tables using `xlsxwriter`.

## Installation

```
pip install git+git://github.com/sztal/xlsxpandas.git
```

## Process

1. Drawer moves over a worksheet and places elements on it.
2. There is one class for atomic elements (single or single merged cells) called `Element`.
   It stores its value, styling (as dict that is later transformed to proper style object),
   and additional parameters like width/height for merged cells or
   comment text and additional comment params.
3. There are several complex elements that are composed of multiple Element objects.
   `Series` and `DataFrame` classes implement corresponding `pandas` classes as
   drawable elements that can be fed into a drawer object and placed within an excel worksheet.
   `Dictionary` object defines a set of key-value pairs.
   Internally it is represented as `list` of `dict`s or `OrderedDict`s.

## Usage

The main idea of `xlsxpandas` is to define a so-called drawer object within an
excel worksheet and use it to draw various drawing elements such as series,
data frames or dictionaries (sets of key-value pairs).

Drawer is initialized in an arbitrary place on a worksheet and then it is
provided with element objects and it places them in the worksheet according
to its own current position. The upper-left cornet of an element is always placed
in the drawer's position.

### Basic examples

Below are several simple examples showing the basic functionalities of `xlsxpandas`.

```python
# Two simplest possible cases: single cell and single merged cell

import xlsxwriter
from xlsxpandas import drawer
from xlsxpandas.elements import Element

wb = xlsxwriter.workbook.Workbook('basic-example.xlsx')
ws = wb.add_worksheet()
dr = drawer.Drawer(ws, wb)  # by default drawer is placed in the first cell (0, 0) / A1

single_cell = Element('single cell')
merged_cell = Element('merged cell', width = 3, height = 2)
single_cell2 = Element(2)

dr.draw(single_cell)
dr.move_vertical()      # move vertically over the last drawn element
dr.draw(merged_cell)
dr.move_horizontal()    # move horizontally over the last drawn element
dr.draw(single_cell2)

wb.close()
```

#### Styling

It is also very easy to add styling to elements.
In `xlsxpandas` elements' styles are always defined as `dict`s.
Only when an element draws itself in a worksheet, the style is definition
is transformed to a proper `xlsxwriter` style object.
This enables easy tranforming and merging of styles between cells before drawing
(i.e. via `{**dict1, **dict2}` syntax).

```python
# Simple styling examples

import xlsxwriter
from xlsxpandas import drawer
from xlsxpandas.elements import Element

wb = xlsxwriter.workbook.Workbook('basic-styling.xlsx')
ws = wb.add_worksheet()
dr = drawer.Drawer(ws, wb)

single_cell = Element('single cell', style = {'bg_color': 'blue'})
merged_cell = Element('merged cell', style = {'border': 2, 'bold': True},
                      width = 3, height = 2)

dr.draw(single_cell)
dr.move_horizontal()
dr.draw(merged_cell)

wb.close()
```

#### Comments and `xlsxwriter` write methods

Adding and customizing comments as well as using arbitrary `xlsxwriter`
write methods is also supported by `xlsxpandas`.

```python
# Comments and write methods

import xlsxwriter
from xlsxpandas import drawer
from xlsxpandas.elements import Element

wb = xlsxwriter.workbook.Workbook('comments-and-write-methods.xlsx')
ws = wb.add_worksheet()
dr = drawer.Drawer(ws, wb)

cell = Element('https://www.google.com', width = 2, height = 2,
               comment = 'a link', comment_params = {'author': 'an author'},
               write_method = 'write_url', write_args = {'string': 'a link'})
dr.draw(cell)

wb.close()
```

#### Grid of elements

This above examples already shows an advantage of using `xlsxpandas` over 
sole `xlsxwriter`. By having a plotter-like drawer object it is very easy
to keep track of what and where is being drawn and methods like `move_horizontal`
and `move_vertical` that goes over the las drawn element make it possible
to move easily throughout a worksheet without any manual bookkeping of coordinates.

This is easily shown in the example below where a simple `dict` is drawn
as a raggd table with header.

```python
# Ragged table of elements

import xlsxwriter
from xlsxpandas import drawer
from xlsxpandas.elements import Element

wb = xlsxwriter.workbook.Workbook('simple-ragged-table.xlsx')
ws = wb.add_worksheet()
dr = drawer.Drawer(ws, wb)

data = {
    'A': range(10),
    'B': range(6),
    'Rather Quite a Long Name': range(12)
}

start_x = dr.x
for key, value in data.items():
    head = Element(key, style = {'bold': True, 'align': 'center'},
                   col_width = 'auto', padding = 2)
    
    # col_width = 'auto' make autoadjustment of the column accordingly to the
    # width of its content + adds padding on both sides
    # col_width argument may also take a specific width given as a flot.
    
    dr.draw(head)
    dr.move_vertical()
    for val in value:
        elem = Element(val)
        dr.draw(elem)
        dr.move_vertical()
    dr.reset(x = start_x, y = None) # reset x-coordinate and leave y-cooridnate untouched)
    dr.move_horizontal()

wb.close()
```

But what is even better, operations like this are even more simple thanks to
`xlsxpandas` extensions of `pandas` classes like `Series` and `DataFrame`
as well as custom complex elements like `Dictionary`.

### Series

The `Series` class is an extension of the corresponding `pandas` class and can
be used to easily draw vertical or horizontal lines of elements.
It also supports adding arbitrary styeles to first and last cells,
so fo instance it is easy to properly border entire region.

```python
# Series usage examples

import xlsxwriter
import pandas as pd
from xlsxpandas import drawer
from xlsxpandas.elements import Series

wb = xlsxwriter.workbook.Workbook('series.xlsx')
ws = wb.add_worksheet()
dr = drawer.Drawer(ws, wb)

# Series with borders
s1 = Series(range(15), borders = 2)
# Horizontal series (i.e. a table's header)
s2 = Series(['A', 'B', 'C', 'D'], borders = 2, style = {'bold': True},
               horizontal = True)
dr.draw(s1)
dr.move_horizontal()
dr.draw(s2)
dr.move_vertical()
for i in range(2):
    s = Series(range(14), borders = 1)
    dr.draw(s)
    dr.move_horizontal()

# Series of width 2 and numbers written as text
# xlsxpandas series may be initialized from any valid `data` value for pandas Series,
# including other pandas series.
s = pd.Series(range(14)).astype(str)
s = Series(s, width = 2, borders = 1, write_method = 'write_string',
           style = {'align': 'center'})

dr.draw(s)

wb.close()
```

### DataFrame

The `DataFrame` class is an extension of the `pandas` class and can be used
to easily draw entire tables. It can be provided with arbitrary column specifications
via `col_args` (it stores arguments passed to constructors of given columns),
as well as table-level settings including styling (and border styling)
and write_methods with arbitrary additional parameters.

```python
# DataFrame examples

import xlsxwriter
import pandas as pd
from xlsxpandas import drawer
from xlsxpandas.elements import DataFrame, Series

wb = xlsxwriter.workbook.Workbook('data-frame.xlsx')
ws = wb.add_worksheet()
dr = drawer.Drawer(ws, wb)

# It is often most handy to build data frames from dicts
df = DataFrame(
    {
        'A': [1, 2, 3],
        'Long Column Name': ['aaaaaaaaaaaaaaa', 'bbb', 'CCC'],
        'URL': ['https://www.google.com', 'https://www.google.com', 'https://www.google.com']
    },
    borders = 2,
    col_args = {
        'Long Column Name': {
            'col_width': 'auto',
            'style': {'bg_color': 'gray'}
        }
    }
)

# URL strings can be set using indexing methods
df.loc[:, 'URL'] = df.loc[:, 'URL'] \
                     .pipe(Series) \
                     .setprop('write_method', 'write_url') \
                     .setprop('write_args', [{'string': 'Link1'}, 
                                             {'string': 'Link2'}, 
                                             {'string': 'Link3'}])
dr.draw(df)

wb.close()
```

### Dictionary

The `Dictionary` class is an implementation of key-value fieldsets.
It is useful for example when one need draw a summary or intro page in a report.
`Dictionary` objects are internally represented as list of `dict`s or `OrderedDict`s.
Every dict needs to have a `key` and `value` field, which is an atomic value
or anothr key-value map that specifies proprties of the item at hand.

`Dictionary` class is also equipped with a method (`load_config`) for 
reading `.yaml` files using `OrderedDict` representation.
It can be useful, so it is often easier to define the structure of
a complex `Dictionary` object in a separate `.yaml` file intead of a code.

```python
# Dictionary example

import xlsxwriter
import pandas as pd
from xlsxpandas import drawer
from xlsxpandas.elements import Dictionary

wb = xlsxwriter.workbook.Workbook('dictionary.xlsx')
ws = wb.add_worksheet()
dr = drawer.Drawer(ws, wb)

dictionary = Dictionary(
    [
        {
            'key': {
                'value': 'key1'
            },
            'value': {
                'value': ['value1', 'value2']
            }
        },
        {
            'key': {
                'value': 'key2'
            },
            'value': {
                'value': 'value3',
                'style': {
                    'color': 'blue',
                    'bg_color': 'yellow',
                    'bold': True
                }
            }
        }
    ],
    vspace = 1,
    hspace = 2,
    keys_params = {
        'italic': True
    },
    values_params = {
        'underline': True
    }
)

dr.draw(dictionary)

wb.close()
```
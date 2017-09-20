# -*- coding: utf-8 -*-
"""Drawing element classes"""

# Imported modules ------------------------------------------------------------

# Full imports ---
import sys, yaml, re
import pandas as pd

# Partial imports ----
from collections import OrderedDict
from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxpandas.__internals__ import (
    validate_param
)

###############################################################################

class Element(object):
    """Implementation of an atomic report element.
    
    It is fed to a drawer object and then
    it is drawn in the supplied matrix xy-coordinates on a worksheet
    on which the drawer is located.
    If height or width is greater than 1, then appropriate cells
    (counting from the top-left corner are merged).
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def value(self):
        """any atomic: displayed valued of the element.
        """
        return self._value
    @value.setter
    def value(self, value):
        self._value = value
    
    @property
    def height(self):
        """positive int: height of the element in worksheet cells."""
        return self._height
    @height.setter
    def height(self, value):
        self._height = validate_param(value, 'height', int, True, 'x > 0')
    
    @property
    def width(self):
        """positive int: width of the element int worksheet cells."""
        return self._width
    @width.setter
    def width(self, value):
        self._width = validate_param(value, 'width', int, True, 'x > 0')
    
    @property
    def style(self):
        """dict: `xlsxwriter`-compatible style definitions for the element."""
        return self._style
    @style.setter
    def style(self, value):
        self._style = validate_param(value, 'style', dict)
    
    @property
    def comment(self):
        """str: comment text (optional)."""
        return self._comment
    @comment.setter
    def comment(self, value):
        self._comment = validate_param(value, 'comment', (str, type(None)))
    
    @property
    def comment_params(self):
        """dict: additional comment parameters (optional)."""
        return self._comment_params
    @comment_params.setter
    def comment_params(self, value):
        self._comment_params = validate_param(value, 'comment_params', dict)
    
    @property
    def write_method(self):
        """str: name of a `xlsxwriter.worksheet` write method."""
        return self._write_method
    @write_method.setter
    def write_method(self, value):
        self._write_method = validate_param(value, 'write_method', str)
    
    @property
    def write_args(self):
        """dict: additional arguments passed to the `write_method` (optional)."""
        return self._write_args
    @write_args.setter
    def write_args(self, value):
        self._write_args = validate_param(value, 'write_args', dict)
    
    @property
    def col_width(self):
        """float, 'auto' or None: width settings for columns spaned by the element.
        If 'auto' then with is determined according to the width of the content + padding.
        """
        return self._col_width
    @col_width.setter
    def col_width(self, value):
        self._col_width = \
            validate_param(value, 'col_width', (float, str, type(None)),
                           lambda x: x if isinstance(x, (str, type(None))) else float(x),
                           'x > 0 if isinstance(x, float) else True')
    
    @property
    def padding(self):
        """float: additional padding added to both sides if `col_width = 'auto'` (optional)."""
        return self._padding
    @padding.setter
    def padding(self, value):
        self._padding = validate_param(value, 'padding', float, True, 'x >= 0')
    
    # -------------------------------------------------------------------------
    
    def __init__(self, value, height = 1, width = 1, style = {}, 
                 comment = None, comment_params = {},
                 write_method = 'write', write_args = {},
                 col_width = None, padding = 2.0):
        """Initilization method.
        
        Parameters
        ----------
        value : any atomic type or a tuple
            Value to be written in the element.
        height : int
            Height as a number of cells (rows); non-negative.
        width : int
            Width as a number of cells (columns); non-negative.
        style : dict
            Element's style definition
        comment : str
            Comment text; defaults to `None`.
        comment_params : dict
            Comment params (see [xlsxwriters docs](xlsxwriter.readthedocs.io/worksheet.html#worksheet-write-comment)); defaults to {}.
        write_method : str
            Name of a `xlsxwriter.worksheet` write method to use; 
            defaults to generic `write`;
            should be one of the valid `xlsxwriter.worksheet` write methods 
            (including `write_rich_string`).
        write_args : dict
            Optional keyword arguments passed to the write method.
        col_width : float, 'auto' or None
            Width of the column. If the element's width is greater than 1, then width determines total width of all columns.
        padding : float
            Padding added on both sides when `col_width = 'auto'`.
        """
        self.value = value
        self.height = height
        self.width = width
        self.style = style
        self.comment = comment
        self.comment_params = comment_params
        self.write_method = write_method
        self.write_args = write_args
        self.col_width = col_width
        self.padding = padding
    
    def _make_style(self, wb):
        """Register Element's style for drawing"""
        return wb.add_format(self.style)
    
    def xl_upleft(self, x, y):
        """Get upper-left corner coordinates of the Element 
        in the standard excel notation.
        
        Parameters
        ----------
            x : int
                Element's x-coordinate (its upper-left corner).
            y : int
                Element's y-coordinate (its upper-left corner).
        """
        return xl_rowcol_to_cell(x, y)
    
    def xl_loright(self, x, y):
        """Get lower-right corner cooridnates of the Element 
        in the standard excel notation.
        
        Parameters
        ----------
            x : int
                Element's x-coordinate (its upper-left corner).
            y : int
                Element's y-coordinate (its upper-left corner).
        """
        return xl_rowcol_to_cell(x + self.height - 1, y + self.width - 1)
    
    def xl_range(self, x, y):
        """Get range covered with the Element in the standard excel notation
        
        Parameters
        ----------
            x : int
                Element's x-coordinate (its upper-left corner).
            y : int
                Element's y-coordinate (its upper-left corner).
        """
        upleft = self.xl_upleft(x, y)
        loright = self.xl_loright(x, y)
        return upleft + ':' + loright
    
    def draw(self, x, y, ws, wb, na_rep, **kwargs):
        """Draw Element in the worksheet.
        
        This method is public, but usually should not be used 'by hand'.
        Instead, it is called by a drawer object's method `draw`.
        
        Parameters
        ----------
            x : int
                X-coordinate for the upper-left corner of the Element.
            y : int
                Y-coordinate for the upper-left corner of the Element.
            ws : xlsxwriter.worksheet.Worksheet
                Worksheet to write the Element in.
            wb : xlsxwriter.workbook.Workbook
                Workbook the worksheet is in.
            na_rep : str
                String representation of missing values.
            **kwargs : any
                Optional keyword parameters passed to the write methods.
        """
        if self.write_method == 'write_rich_string' and isinstance(self.value, tuple):
            vals = list(self.value)
            vals = [ wb.add_format({**self.style, **x}) if isinstance(x, dict) else x for x in vals ]
            self.value = tuple(vals)
        wmethod = getattr(ws, self.write_method)
        style = self._make_style(wb)
        wargs = {**self.write_args, **kwargs}
        if isinstance(self.value, tuple) and self.write_method == 'write_rich_string':
            if self.height > 1 or self.width > 1:
                ws.merge_range(self.xl_range(x, y), '', style)
            wmethod(self.xl_upleft(x, y), *self.value)
        else:
            if self.width > 1 or self.height > 1:
                rng = self.xl_range(x, y)
                ws.merge_range(rng, '', style)
            wmethod(x, y, self.value, style, **wargs)
            if self.comment is not None:
                addr = self.xl_upleft(x, y)
                ws.write_comment(addr, self.comment, self.comment_params)
        
        # Apply column width adjustment
        def vlen(value):
            if value is not None:
                return len(str(value))
            else:
                return None
        
        if hasattr(self, 'col_width'):
            if isinstance(self.col_width, float):
                col_width = self.col_width
            elif isinstance(self.col_width, str) and self.col_width == 'auto':
                try:
                    col_width = float(vlen(self.value) + self.padding * 2) / self.width
                except TypeError:
                    return
            elif self.col_width is None:
                return
            else:
                raise ValueError('incorrect value of col_width.')
            ws.set_column(y, y + self.width - 1, col_width / self.width)

###############################################################################

class Series(pd.Series):
    """Series of elements
    
    This class utilizes functionalities of pandas.Series class.
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def horizontal(self):
        """bool: Should the series be layed out horizontally or vertically.
        Defaults to to `True`.
        """
        return self._horizontal
    @horizontal.setter
    def horizontal(self, value):
        self._horizontal = validate_param(value, 'horizontal', bool)
    
    @property
    def length(self):
        """int: Length of the series (read-only)."""
        if self.horizontal:
            return sum([ x.width for x in self.values ])
        else:
            return sum([ x.height for x in self.values ])
    
    @property
    def width(self):
        """int: Width of the series (read-only)."""
        if self.horizontal:
            return self.length
        else:
            return self.apply(lambda x: x.width).max()
    
    @property
    def height(self):
        """int: Height of the series (read-only)."""
        if self.horizontal:
            return self.apply(lambda x: x.height).max()
        else:
            return self.length
    
    @property
    def col_width(self):
        """float, 'auto' or None: width settings for columns spanned by the series.
        If 'auto' then with is determined according to the width of the content + padding.
        """
        return self._col_width
    @col_width.setter
    def col_width(self, value):
        self._col_width = \
            validate_param(value, 'col_width', (float, str, type(None)),
                           lambda x: x if isinstance(x, (str, type(None))) else float(x),
                           'x > 0 if isinstance(x, float) else True')
    
    @property
    def padding(self):
        """float: additional padding added to both sides if `col_width = 'auto'` (optional)."""
        return self._padding
    @padding.setter
    def padding(self, value):
        self._padding = validate_param(value, 'padding', float, True, 'x >= 0')
    
    # -------------------------------------------------------------------------
    
    def __init__(self, data, horizontal = False, height = 1, width = 1,
                 style = {}, name_args = {}, 
                 borders = None, first = {}, last = {}, 
                 write_method = 'write', write_args = {}, 
                 col_width = None, padding = 2.0, **kwargs):
        """Initialization method.
        
        Parameters
        ----------
        data : array-like, dict or scalar value
            Elements that series is to be made of.
        horizontal : bool
            Should series be aligned horizontally or vertically.
        height : int
            Height of cells of the series.
        width : int
            Width of cells of the series.
        style : dict
            Base style of the series.
        name_args : dict
            Additional arguments passed to name element constructor.
        borders : int or None
            Borders definition for the series; 
            set side as well as first and last elements' borders.
        first : int or dict
            Additional styling for the first element of the series;
            if `int` then it is a border value; 
            `dict` is an arbitrary styling compatible with `xlsxwriter`.
        last : int or dict
            Additional styling for the last element of the series.
        write_method : str
            Name of a `xlsxwriter.worksheet` write method for the series.
        write_args : dict
            Additional arguments for the write method of the series.
        col_width : float, 'auto' or None
            Width of the column. If the element's width is greater than 1, 
            then width determines total width of all columns.
        padding : float
            Padding added on both sides when `col_width = 'auto'`.
        **kwargs 
            Other optional parameters passed to the pandas Series constructor.
        """
        super(Series, self).__init__(data, **kwargs)
        
        side1 = 'top' if horizontal else 'left'
        side2 = 'bottom' if horizontal else 'right'
        fpos = 'left' if horizontal else 'top'
        lpos = 'right' if horizontal else 'bottom'
        if borders:
            style[side1] = borders
            style[side2] = borders
            first = {fpos: borders}
            last = {lpos: borders}
        
        # Initilize elements ---
        for i in self.index:
            if not isinstance(self[i], Element):
                if isinstance(self[i], (dict, OrderedDict)):
                    elem = Element(**self[i])
                    elem.style = {**style, **elem.style}
                else:
                    elem = Element(self[i], height, width, style,
                                   write_method = write_method, write_args = write_args)
                self[i] = elem
        
        # Determine name element ---
        if self.name:
            if not isinstance(self.name, Element):
                if isinstance(self.name, (dict, OrderedDict)):
                    self.name = Element(**self.name)
                    self.name.style = {**style, **self.name.style}
                else:
                    stl = name_args.pop('style', {})
                    self.name = Element(self.name, height, width,
                                        {**style, **stl}, **name_args)
        
        # Determine first and last elements' styles ---
        felem = self.iloc[0]
        lelem = self.iloc[-1]
        fstl = {**felem.style, fpos: first} if isinstance(first, int) \
                                            else {**felem.style, **first}
        lstl = {**lelem.style, lpos: last} if isinstance(last, int) \
                                           else {**lelem.style, **last}
        felem.style = fstl
        lelem.style = lstl
        self.iloc[0] = felem
        self.iloc[-1] = lelem         
        
        self.horizontal = horizontal
        self.col_width = col_width
        self.padding = padding
    
    def setprop(self, propname, value, inplace = False):
        """Set a property of all elements in the series.
        
        It is useful because it may be used after flitering the series
        with the `pandas` methods like `loc` or `iloc`.
        
        Parameters
        ----------
            propname : str
                Property name.
            value : any or a list with the same length as the series
                New value.
            inplace : bool
                Should assignment be done in place; defaults to `False`.
        """
        sr = self if inplace else self.copy()
        if isinstance(value, list):
            if len(value) != sr.size:
                raise ValueError('`value` has different length than the series.')
            for i, val in zip(sr.index, value):
                setattr(sr[i], propname, val)
        else:
            for i in sr.index:
                setattr(sr[i], propname, value)
        if not inplace:
            return self
    
    def addstyle(self, style, inplace = False):
        """Add additional styling to the existing style.
        
        For overwriting styles the `setprop` method should be used.
        
        Parameters
        ----------
            style : dict or list of dicts with the same length as the series
                Additional styling definitions.
            inplace : bool
                Should assignment be done in place; defaults to `False`.
        """
        sr = self if inplace else self.copy()
        if isinstance(style, list):
            if len(style) != sr.size:
                raise ValueError('`style` has differen length than the series.')
            for i, stl in zip(sr.index, style):
                sr[i].style = {**sr[i].style, **stl}
        else:
            for i in sr.index:
                sr[i].style = {**sr[i].style, **style}
        if not inplace:
            return self
    
    def draw(self, x, y, ws, wb, na_rep, draw_name = False, **kwargs):
        """Draw Series in the worksheet.
        
        Parameters
        ----------
            x : int
                X-coordinate for the upper-left corner of the Series.
            y : int
                Y-coordinate for the upper-left corner of the Series.
            ws : xlsxwriter.worksheet.Worksheet
                Worksheet to write the Element in.
            wb : xlsxwriter.workbook.Workbook
                Workbook the worksheet is in.
            na_rep : str
                String representation of missing values.
            draw_name : bool
                Should name element be drawn (if defined).
            **kwargs
                Optional keyword parameters passed to the write methods.
        """
        if self.horizontal:
            if draw_name and self.name:
                self.name.draw(x, y, ws, wb, na_rep, **kwargs)
                y += self.name.width
            for elem in self.values:
                elem.draw(x, y, ws, wb, na_rep, **kwargs)
                y += elem.width
        else:
            if draw_name and self.name:
                self.name.draw(x, y, ws, wb, na_rep, **kwargs)
                x += self.name.height
            for elem in self.values:
                elem.draw(x, y, ws, wb, na_rep, **kwargs)
                x += elem.height
        
        # Apply column width adjustment
        def vlen(value):
            if value is not None:
                return len(str(value))
            else:
                return None
        
        if self.col_width and not self.horizontal:
            if isinstance(self.col_width, (float, int)):
                col_width = float(self.col_width)
            elif isinstance(self.col_width, str) and self.col_width == 'auto':
                try:
                    col_width = max([ vlen(elem.value) for elem in self ])
                    col_width = (col_width + self.padding * 2) / self.width
                except TypeError:
                    return
            elif self.col_width is None:
                return
            else:
                raise ValueError('incorrect value of col_width.')
            ws.set_column(y, y + self.width - 1, col_width / self.width)
        elif self.col_width and self.horizontal:
            if isinstance(self.col_width, str) or self.col_width is None:
                return
            elif isinstance(self.col_width, (int, float)):
                row_width = float(self.col_width)
                for i in range(self.height):
                    ws.set_row(x + i, row_width / self.height)            
        
###############################################################################

class DataFrame(pd.DataFrame):
    """DataFrame of elements.
    
    This class utilizes functionalities of pandas.DataFrame class.
    Current implementation does not support data frames with hierarchical indexes,
    and it may not work correctly for such cases.
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def width(self):
        """positive int: width of the data frame."""
        return self.apply(lambda x: sum([ y.width for y in x ]), axis = 1).max()
    
    @property
    def height(self):
        """positive int: height of the data frame."""
        return self.apply(lambda x: sum([ y.height for y in x ]), axis = 0).max()
    
    @property
    def name_args(self):
        """dict: additional arguments for name elements constructors."""
        return self._name_args
    @name_args.setter
    def name_args(self, value):
        self._name_args = validate_param(value, 'name_args', dict)
    
    @property
    def col_args(self):
        """dict: additional arguments for columns constructors."""
        return self._col_args
    @col_args.setter
    def col_args(self, value):
        self._col_args = validate_param(value, 'col_args', dict)
    
    # -------------------------------------------------------------------------
    
    def __init__(self, data, height = 1, width = 1, style = {}, 
                 borders = None, top = {}, bottom = {}, left = {}, right = {}, 
                 write_method = 'write', write_args = {},
                 name_args = {}, col_args = {}, **kwargs):
        """Initialization method.
        
        Parameters
        ----------
        data : numpy ndarray (structured or homogeneous), dict, or DataFrame
            Elements that data frame is to be made of.
        height : int
            Height of the elements of the data frame.
        width : int
            Width of the elements of the data frame.
        style : dict
            Base style of the data frame.
        borders : int or None
            Set border defintitions for boundary elements.
        top : int or dict
            Additional styling for the top row of the data frame;
            if int then it is a border value; 
            dict is an arbitrary styling compatible with `xlsxwriter`.
        bottom : int or dict
            Additional styling for the bottom row of the data frame.
        left : int or dict
            Additional styling for the leftmost column of the data frame.
        right : int or dict
            Additional styling for the rightmost column of the data frame.
        write_method : str
            Write method of the data frame. 
            One of the `xlsxwriter.worksheet` write methods.
        write_args : dict
            Addtional arguments for the write method of the data frame.
        name_args : dict
            Additional arguments passed to name element constructors if names are drawn.
        col_args : dict
            Additional arbitrary arguments passed to column series constructor while drawing.
        **kwargs
            other optional parameters passed to the pandas DataFrame constructor
        """
        super(DataFrame, self).__init__(data, **kwargs)
        
        if borders:
            top = {**top, 'top': borders}
            right = {**right, 'right': borders}
            bottom = {**bottom, 'bottom': borders}
            left = {**left, 'left': borders}
        
        # Initialize elements ---
        for i, row in self.iterrows():
            for j, elem in row.iteritems():
                if not isinstance(elem, Element):
                    if isinstance(elem, dict):
                        stl = elem.pop('style', {})
                        elem = Element(**elem, style = {**style, **stl})
                    else:
                        elem = Element(elem, height, width, style,
                                       write_method = write_method,
                                       write_args = write_args)
                    try:
                        self.iloc[i, j] = elem
                    except ValueError:
                        self.loc[i, j] = elem
                
        # Determine boundary styles ---
        top = {'top': top} if isinstance(top, int) else top
        bottom = {'bottom': bottom} if isinstance(bottom, int) else bottom
        left = {'left': left} if isinstance(left, int) else left
        right = {'right': right} if isinstance(right, int) else right
        
        # Apply boundary styles ---
        for i in range(self.shape[1]):
            self.iloc[0, i].style = {**self.iloc[0, i].style, **top}
            self.iloc[-1, i].style = {**self.iloc[-1, i].style, **bottom}
        for i in range(self.shape[0]):
            self.iloc[i, 0].style = {**self.iloc[i, 0].style, **left}
            self.iloc[i, -1].style = {**self.iloc[i, -1].style, **right}
        
        self.col_args  = col_args
        self.name_args = name_args
    
    def setprop(self, propname, value, inplace = False):
        """Set a property of all elements in the data frame.
        
        It is useful because it may be used after flitering the data frame.
        It does not support multiple new values.
        It is better to do multiple assignments via specfic series and theirs `setprop` methods.
        
        Parameters
        ----------
            propname : str
                Property name.
            value : any
                New value.
            inplace : bool
                Should assignment be done in place; defaults to `False`.
        """
        df = self if inplace else self.copy()
        for i in df.index:
            for j in df.columns:
                try:
                    setattr(df.iloc[i, j], propname, value)
                except ValueError:
                    setattr(df.loc[i, j], propname, value)
        if not inplace:
            return self
    
    def addstyle(self, style, inplace = False):
        """Add additional styling to the existing style.
        
        For overwriting styles the `setprop` method should be used.
        Multiple style alterations should be done on series' level via `Series.addstyle` method.
        
        Parameters
        ----------
            style : dict
                Additional styling definitions.
            inplace : bool
                Should assignment be done in place; defaults to `False`.
        """
        df = self if inplace else self.copy()
        for i in df.index:
            for j in df.columns:
                try:
                    df.iloc[i, j].style = {**df.iloc[i, j].style, **style}
                except ValueError:
                    df.loc[i, j].style = {**df.loc[i, j].style, **style}
        if not inplace:
            return df
    
    def draw(self, x, y, ws, wb, na_rep, draw_names = False, **kwargs):
        """Draw DataFrame in the worksheet
        
        Parameters
        ----------
            x : int
                X-coordinate for the upper-left corner of the DataFrame.
            y : int
                Y-coordinate for the upper-left corner of the DataFrame.
            ws : xlsxwriter.worksheet.Worksheet
                Worksheet to write the Element in.
            wb : xlsxwriter.workbook.Workbook
                Workbook the worksheet is in.
            na_rep : str
                String representation of missing values.
            draw_names : bool
                Should column names be draw; defaults to `False`.
            **kwargs
                Optional keyword parameters passed to the write methods.
        """
        for index, col in self.iteritems():
            cargs = self.col_args.get(index, {})
            stl   = cargs.pop('style', {})
            nargs = {**self.name_args, **cargs.pop('name_args', {})}
            nargs['style'] = {**nargs.get('style', {}), **stl}
            col = Series(col, name_args = nargs, **cargs) \
                  .addstyle(stl)
            col.draw(x, y, ws, wb, na_rep, draw_name = draw_names, **kwargs)
            y += col.width

###############################################################################
        
class Dictionary(object):
    """Visual/tabular representaion of a key => value set.
    
    This class implements a layout of fields in a report,
    in which there is one column (a key column)
    separated from a second column by a horizontal space of a given width
    that presents key (titles) and the second column presents content (values)
    for given keys. Useful form making into/definitions pages for various reports.
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def structure(self):
        """list: Dictionary structure definition."""
        return self._structure
    @structure.setter
    def structure(self, value):
        if isinstance(value, str):
            value = self.load_config(value)
        self._structure = validate_param(value, 'structure', list)
    
    @property
    def hspace(self):
        """nonnegative int: Additional horizontal spaceing between keys and values."""
        return self._hspace
    @hspace.setter
    def hspace(self, value):
        self._hspace = validate_param(value, 'hspace', int, True, 'x >= 0')
    
    @property
    def vspace(self):
        """nonnegative int: Additional vertical spaceing between items."""
        return self._vspace
    @vspace.setter
    def vspace(self, value):
        self._vspace = validate_param(value, 'vspace', int, True, 'x >= 0')
    
    @property
    def keys_params(self):
        """dict: Additional styling for keys."""
        return self._keys_params
    @keys_params.setter
    def keys_params(self, value):
        self._keys_params = validate_param(value, 'keys_params', dict)
    
    @property
    def values_params(self):
        """dict: Additional styling for values."""
        return self._values_params
    @values_params.setter
    def values_params(self, value):
        self._values_params = validate_param(value, 'values_params', dict)
    
    @property
    def context(self):
        """dict: Additional variables used when evaluating expressions
        embedded in keys or values content (`@eval@` syntax).
        """
        return self._context
    @context.setter
    def context(self, value):
        self._context = validate_param(value, 'context', dict)
    
    @property
    def width(self):
        """positive int: Width of the dictionary."""
        width = 0
        for elem in self.structure:
            w = elem['key'].get('width', 1)
            w += elem['value'].get('width', 1)
            w += elem.get('hspace', self.hspace)
            if w > width:
                width = w
        return width
    
    @property
    def height(self):
        """positive int: Height of the dictionary."""
        height = 0
        for elem in self.structure:
            w = elem['key'].get('height', 1)
            if isinstance(elem['value']['value'], list):
                try:
                    vw = sum([ e.get('height', 1) for e in elem['value']['value'] ])
                except AttributeError:
                    vw = len(elem['value']['value']) * elem['value'].get('height', 1)
                if vw > w:
                    w = vw
            height += w
            height += elem.get('vspace', self.vspace)
        return height
    
    # -------------------------------------------------------------------------
    
    def __init__(self, structure, hspace = 1, vspace = 0,
                 keys_params = {}, values_params = {}, context = {}):
        """Initilization method.
        
        Parameters
        ----------
        structure : OrderedDict or path to a `.yaml` file defining the structure
            Dictionary structure definition.
        hspace : int (>= 0)
            Additional horizontal spacing between keys and values.
        vspace : int (>= 0)
            Additional vertical spacing between elements (may be overwritten by element-level settings).
        keys_params : dict
            Default style definitions for keys.
        values_params : dict
            Default style definitions for params.
        context : dict
            Additional contextual variables that may be used in the structure definition.
            Syntax @eval@'expression' enables evaluation of expressions 
            in the structure using the context variables.
        """
        self.structure = structure
        self.hspace = hspace
        self.vspace = vspace
        self.keys_params = keys_params
        self.values_params = values_params
        self.context = context
    
    @staticmethod
    def load_config(path = None):
        """Loads config from a config.yaml file.
    
        Parameters
        ----------
            path : str
                Path to a config file.
            
        Returns
        -------
            list
                config parsed to a `list` of `OrderedDicts`.
        """            
        def ordered_load(stream, Loader = yaml.Loader, object_pairs_hook = OrderedDict):
            class OrderedLoader(Loader):
                pass
            def construct_mapping(loader, node):
                loader.flatten_mapping(node)
                return object_pairs_hook(loader.construct_pairs(node))
            OrderedLoader.add_constructor(
                yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG,
                construct_mapping
            )
            return yaml.load(stream, OrderedLoader)    
        cnf = open(path, 'r')
        try:
            config = ordered_load(cnf)
        except yaml.YAMLError as exc:
            sys.exit(exc)
        finally:
            cnf.close()
        return config   
    
    def _process_value(self, x):
        """Evaluate string agains a context.
        """
        if isinstance(x, str) and re.match('^@eval@', x):
            x = re.sub('^@eval@', '', x)
            return eval(x, None, self.context)
        else:
            return x

    def draw(self, x, y, ws, wb, na_rep, **kwargs):
        """Draw Dictionary in a worksheet.
        
        Parameters
        ----------
            x : int
                X-coordinate for the upper-left corner of the Dictionary.
            y : int
                Y-coordinate for the upper-left corner of the Dictionary.
            ws : xlsxwriter.worksheet.Worksheet
                Worksheet to write the Element in.
            wb : xlsxwriter.workbook.Workbook
                Workbook the worksheet is in.
            na_rep : str
                String representation of missing values.
            **kwargs
                Optional keyword parameters passed to the write methods.
        """
        y0 = y
        for elem in self.structure:
            elem['key']['value'] = self._process_value(elem['key']['value'])
            elem['key']['style'] = {**self.keys_params, **elem['key'].get('style', {})}
            if isinstance(elem['value']['value'], list):
                elem['value']['value'] = [ self._process_value(x) for x in elem['value']['value'] ]
            else:
                elem['value']['value'] = self._process_value(elem['value']['value'])
            elem['value']['style'] = {**self.values_params, **elem['value'].get('style',{})}
            key = Element(**elem['key'])
            key.draw(x, y, ws, wb, na_rep, **kwargs)
            y += self.hspace + 1
            values = elem['value']['value']
            if isinstance(values, list):
                for value in elem['value']['value']:
                    if isinstance(value, (dict, OrderedDict)):
                        e = Element(**{**elem['value'], **value})
                    else:
                        e = Element(**{**elem['value'], 'value': value})
                    e.draw(x, y, ws, wb, na_rep, **kwargs)
                    x += e.height
                x += elem.get('vspace', self.vspace)
            else:
                e = Element(**{**elem['value'], 'value': values})
                e.draw(x, y, ws, wb, na_rep, **kwargs)
                x += e.height + elem.get('vspace', self.vspace)
            y = y0

###############################################################################

"""Drawing element classes"""

# Imported modules ------------------------------------------------------------

# Full imports ---
import xlsxwriter
import sys, yaml, re
import pandas as pd
import numpy as np

# Partial imports ----
from collections import OrderedDict
from xlsxwriter.utility import xl_rowcol_to_cell
from __internals__ import (
    validate_param
)

###############################################################################

class Element(object):
    """Implementation of an atomic report element
    
    It is fed to a Drawer object and then
    it is drawn in the supplied matrix xy-coordinates.
    If height or width is greater than 1, then appropriate cells
    (counting from the top-left corner are merged).
    
    Parameters
    ----------
        value : any atomic type or a tuple
            value to be written in the element
        height : int
            height as a number of cells (rows); non-negative
        width : int
            width as a number of cells (columns); non-negative
        style : dict
            Element's style definition
        comment : str
            comment text; defaults to `None`
        comment_params : dict
            comment params (see xlsxwriters docs); defaults to {}
        write_method : str
            name of a xlsxwriter write method to use; defaults to generic `write`;
            should be one of the valid xlsxwriter worksheet write methods (including `write_rich_string`)
        write_args : dict
            optional keyword arguments passed to the write method
        col_width : float, str ['auto'] for None
            width of the column. If the element's width is greater than 1, then width determines total width of all columns
        padding : float
            padding added on both sides when `col_width = 'auto'`
    
    Returns
    -------
    **Attributes**:
        * value : element's value
        * height : element's height
        * width : element's widt
        * style : element's style
        * comment : element's comment
        * comment_params : element's comment parameters
        * write_method : name of a write method
        * write_args : optional keyword arguments for the write method
        * col_width : optional column width definition
        * padding : padding added to column width when auto resizing
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def value(self):
        return self._value
    @value.setter
    def value(self, value):
        self._value = value
    
    @property
    def height(self):
        return self._height
    @height.setter
    def height(self, value):
        self._height = validate_param(value, 'height', int, True, 'x >= 0')
    
    @property
    def width(self):
        return self._width
    @width.setter
    def width(self, value):
        self._width = validate_param(value, 'width', int, True, 'x >= 0')
    
    @property
    def style(self):
        return self._style
    @style.setter
    def style(self, value):
        self._style = validate_param(value, 'style', dict)
    
    @property
    def comment(self):
        return self._comment
    @comment.setter
    def comment(self, value):
        self._comment = validate_param(value, 'comment', (str, type(None)))
    
    @property
    def comment_params(self):
        return self._comment_params
    @comment_params.setter
    def comment_params(self, value):
        self._comment_params = validate_param(value, 'comment_params', dict)
    
    @property
    def write_method(self):
        return self._write_method
    @write_method.setter
    def write_method(self, value):
        self._write_method = validate_param(value, 'write_method', str)
    
    @property
    def write_args(self):
        return self._write_args
    @write_args.setter
    def write_args(self, value):
        self._write_args = validate_param(value, 'write_args', dict)
    
    @property
    def col_width(self):
        return self._col_width
    @col_width.setter
    def col_width(self, value):
        self._col_width = \
            validate_param(value, 'col_width', (float, str, type(None)),
                           lambda x: x if isinstance(x, (str, type(None))) else float(x),
                           'if isinstance(x, float) x > 0 else True')
    
    @property
    def padding(self):
        return self._padding
    @padding.setter
    def padding(self, value):
        self._padding = validate_param(value, 'padding', float, True, 'x > 0')
    
    # -------------------------------------------------------------------------
    
    def __init__(self, value, height = 1, width = 1, style = {}, 
                 comment = None, comment_params = {},
                 write_method = 'write', write_args = {}):
        """Initilization method
        """
        self.value = value
        self.height = height
        self.width = width
        self.style = style
        self.comment = comment
        self.comment_params = comment_params
        self.write_method = write_method
        self.write_args = write_args
    
    def make_style(self, wb):
        """Register Element's style for drawing
        
        Parameters
        ----------
            wb : xlsxwriter.workbook.Workbook
                workbook to register the style in
        """
        return wb.add_format(self.style)
    
    def xl_upleft(self, x, y):
        """Get upper-left corner coordinates of the Element in the standard excel notation
        
        Parameters
        ----------
            x : int
                element's x-coordinate (its upper-left corner)
            y : int
                element's y-coordinate (its upper-left corner)
        """
        return xl_rowcol_to_cell(x, y)
    
    def xl_loright(self, x, y):
        """Get lower-right corner cooridnates of the Element in the standard excel notation
        
        Parameters
        ----------
            x : int
                element's x-coordinate (its upper-left corner)
            y : int
                element's y-coordinate (its upper-left corner)
        """
        return xl_rowcol_to_cell(x + self.height - 1, y + self.width - 1)
    
    def xl_range(self, x, y):
        """Get range covered with the Element in the standard excel notation
        
        Parameters
        ----------
            x : int
                element's x-coordinate (its upper-left corner)
            y : int
                element's y-coordinate (its upper-left corner)
        """
        upleft = self.xl_upleft(x, y)
        loright = self.xl_loright(x, y)
        return upleft + ':' + loright
    
    def draw(self, x, y, ws, wb, na_rep, **kwargs):
        """Draw Element in the worksheet
        
        Parameters
        ----------
            x : int
                x-coordinate for the upper-left corner of the Element
            y : int
                y-coordinate for the upper-left corner of the Element
            ws : xlsxwriter.worksheet.Worksheet
                worksheet to write the Element in
            wb : xlsxwriter.workbook.Workbook
                workbook the worksheet is in
            na_rep : str
                string representation of missing values
            **kwargs : any
                optional keyword parameters passed to the write methods
        """
        if self.write_method == 'write_rich_string' and isinstance(self.value, tuple):
            vals = list(self.value)
            vals = [ wb.add_format({**self.style, **x}) if isinstance(x, dict) else x for x in vals ]
            self.value = tuple(vals)
        wmethod = getattr(ws, self.write_method)
        style = self.make_style(wb)
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
    
    Parameters
    ----------
        data : array-like, dict or scalar value
            elements that series is to be made of
        name : number, str, Element or None
            if not None, then top element with the value is added as the Series' header
        horizontal : bool
            should series be aligned horizontally or vertically
        first : int or dict
            additional styling for the first element of the series;
            if int then it is a border value; dict is an arbitrary styling compatible with xlsxwriter
        last : int or dict
            additional styling for the last element of the series
        **kwargs : other optional parameters passed to the pandas Series constructor
    
    Returns
    -------
        * all attributes inherited from the pandas Series class
        * name : series name value/element
        * horizontal : series alignment flag
        * length : total length of all elements along the alignment axis
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def horizontal(self):
        return self._horizontal
    @horizontal.setter
    def horizontal(self, value):
        self._horizontal = validate_param(value, 'horizontal', bool)
    
    @property
    def length(self):
        if self.horizontal:
            return sum([ x.width for x in self.values ])
        else:
            return sum([ x.height for x in self.values ])
    @length.setter
    def length(self, value):
        raise AttributeError('length is read-only.')
    
    @property
    def width(self):
        if self.horizontal:
            return self.length
        else:
            return self.apply(lambda x: x.width).max()
    @width.setter
    def width(self, value):
        raise AttributeError('width is read-only.')
    
    @property
    def height(self):
        if self.horizontal:
            return self.apply(lambda x: x.heigth).max()
        else:
            return self.length
    @height.setter
    def height(self, value):
        raise AttributeError('height is read-only.')
    
    # -------------------------------------------------------------------------
    
    def __init__(self, data, name = None, width = 1, horizontal = False,
                 first = {}, last = {}, **kwargs):
        """Initialization method
        """
        super(Series, self).__init__(data, **kwargs)
        
        # Determine first and last elements' styles ---
        felem = self.values[0]
        lelem = self.values[-1]
        fpos = 'left' if horizontal else 'top'
        lpos = 'right' if horizontal else 'bottom'
        fstl = {**felem.style, fpos: first} if isinstance(first, int) \
                                            else {**felem.style, **first}
        lstl = {**lelem.style, lpos: last} if isinstance(last, int) \
                                           else {**felem.style, **last}
        felem.style = fstl
        lelem.style = lstl
        self.values[0] = felem
        self.values[-1] = lelem
        
        self.width = width
        self.horizontal = horizontal
        
        # Determine name ---
        if name:
            if isinstance(name, Element):
                self.name = name
            else:
                self.name = Element(name)
        else:
            self.name = None
        attr = 'height' if horizontal else 'width'
        
        # Determine widths / heights
        if self.name:
            if isinstance(width, int):
                setattr(self.name, attr, width)
            else:
                setattr(self.name, attr, width[0])
        if isinstance(width, int):
            for index, elem in self.iteritems():
                setattr(elem, attr, width)
        else:
            for i in range(len(self.values)):
                setattr(self.values[i], attr, width[i])
    
    def draw(self, x, y, ws, wb, na_rep, **kwargs):
        """Draw Series in the worksheet
        
        Parameters
        ----------
            x : int
                x-coordinate for the upper-left corner of the Series
            y : int
                y-coordinate for the upper-left corner of the Series
            ws : xlsxwriter.worksheet.Worksheet
                worksheet to write the Element in
            wb : xlsxwriter.workbook.Workbook
                workbook the worksheet is in
            na_rep : str
                string representation of missing values
            **kwargs : any
                optional keyword parameters passed to the write methods
        """
        if self.name:
            self.name.draw(x, y, ws, wb, na_rep, **kwargs)
            if self.horizontal:
                y += self.name.width
            else:
                x += self.name.height
        if self.horizontal:
            for elem in self.values:
                elem.draw(x, y, ws, wb, na_rep, **kwargs)
                y += elem.width
        else:
            for elem in self.values:
                elem.draw(x, y, ws, wb, na_rep, **kwargs)
                x += elem.height
        
###############################################################################

class DataFrame(pd.DataFrame):
    """DataFrame of elements
    
    This class utilizes functionalities of pandas.DataFrame class.
    
    Parameters
    ----------
        data : numpy ndarray (structured or homogeneous), dict, or DataFrame
            elements that data frame is to be made of
        top : int or dict
            additional styling for the top row of the data frame;
            if int then it is a border value; dict is an arbitrary styling compatible with xlsxwriter
        bottom : int or dict
            additional styling for the bottom row of the data frame
        left : int or dict
            additional styling for the leftmost column of the data frame
        right : int or dict
            additional styling for the rightmost column of the data frame
    
    Returns
    -------
        * all attributes inherited from the pandas Series class
        * width : total width of the data frame in the excel sheet
        * height : total height of the data frame in the excel sheet
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def width(self):
        return self.apply(lambda x: sum([ y.width for y in x ])).max(axis = 1)
    @width.setter
    def width(self, value):
        raise AttributeError('width is read-only.')
    
    @property
    def height(self):
        return self.apply(lambda x: sum([ y.height for y in x ])).max(axis = 0)
    @height.setter
    def height(self, value):
        raise AttributeError('height is read-only.')
    
    # -------------------------------------------------------------------------
    
    def __init__(self, data, top = {}, bottom = {},
                 left = {}, right = {}, **kwargs):
        """Initialization method
        """
        super(Series, self).__init__(data, **kwargs)
        
        # Determine boundary styles ---
        top = {'top': top} if isinstance(top, int) else top
        bottom = {'bottom': bottom} if isinstance(bottom, int) else bottom
        left = {'left': left} if isinstance(left, int) else left
        right = {'right': right} if isinstance(right, int) else right
        
        # Apply boundary styles ---
        for elem in self.iloc[0, :].iteritems():
            elem.style = {**elem.style, **top}
        for elem in self.iloc[:, -1].iteritems():
            elem.style = {**elem.style, **right}
        for elem in self.iloc[-1, :].iteritems():
            elem.style = {**elem.style, **bottom}
        for elem in self.iloc[:, 0].iteritems():
            elem.style = {**elem.style, **left}
    
    def draw(self, x, y, ws, wb, na_rep, **kwargs):
        """Draw DataFrame in the worksheet
        
        Parameters
        ----------
            x : int
                x-coordinate for the upper-left corner of the DataFrame
            y : int
                y-coordinate for the upper-left corner of the DataFrame
            ws : xlsxwriter.worksheet.Worksheet
                worksheet to write the Element in
            wb : xlsxwriter.workbook.Workbook
                workbook the worksheet is in
            na_rep : str
                string representation of missing values
            **kwargs : any
                optional keyword parameters passed to the write methods
        """
        start_y = y
        for index, row in self.iterrows():
            h = 0
            for elem in row.iteritems():
                elem.draw(x, y, ws, wb, **kwargs)
                y += elem.width
                if elem.height > h:
                    h = elem.height
            x += h
            y = start_y

###############################################################################
        
class Dictionary(object):
    """Visual/tabular representaion of a key => value set
    
    This class implements a layout of fields in a report,
    in which there is one column (a key column)
    separated from a second column by a horizontal space of a given width
    that presents key (titles) and the second column presents content (values)
    for given keys. Useful form making into/definitions pages for various reports.
    
    Parameters
    ----------
        structure : OrderedDict or path to a `.yaml` file defining the structure
            dictionary structure definition
        hspace : int (>= 0)
            additional horizontal spacing between keys and values
        vspace : int (>= 0)
            additional vertical spacing between elements (may be overwritten by element-level settings)
        keys_params : dict
            default set of params passed to Element constructor for the keys
        values_params : dict
            default set of params passed to Element consructor for the values
        context : dict
            additional contextual variables that may be used in the structure definition;
            syntax @eval@'expression' enables evaluation of expressions in the structure using the context variables
    
    Returns
    -------
        * structure : definition of the structure of a Dictionary (key => value) or a path to the .yaml config file
        * hspace : width of the additional horizontal space between key column and value column
        * vspace : additional vertical spacing between elements
        * keys_params : default set of params for key fields
        * values_params : default set of params for values fields
        * context : additional context for evaluation of values
        * width: total width of the dictionary
        * height : total height of the dictionary
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def structure(self):
        return self._structure
    @structure.setter
    def structure(self, value):
        if isinstance(value, str):
            value = self.load_config(value)
        self._structure = validate_param(value, 'structure', OrderedDict)
    
    @property
    def hspace(self):
        return self._hspace
    @hspace.setter
    def hspace(self, value):
        self._hspace = validate_param(value, 'hspace', int, True, 'x >= 0')
    
    @property
    def vspace(self):
        return self._vspace
    @vspace.setter
    def vspace(self, value):
        self._vspace = validate_param(value, 'vspace', int, True, 'x >= 0')
    
    @property
    def keys_params(self):
        return self._keys_params
    @keys_params.setter
    def keys_params(self, value):
        self._keys_params = validate_param(value, 'keys_params', dict)
    
    @property
    def values_params(self):
        return self._values_params
    @values_params.setter
    def values_params(self, value):
        self._values_params = validate_param(value, 'values_params', dict)
    
    @property
    def context(self):
        return self._context
    @context.setter
    def context(self, value):
        self._context = validate_param(value, 'context', dict)
    
    # -------------------------------------------------------------------------
    
    def __init__(self, structure, hspace = 1, vspace = 0,
                 field_params = {}, content_params = {}, context = None):
        """Constructor method
        """
        self.structure = structure
        self.hspace = hspace
        self.vspace = vspace
        self.field_params = field_params
        if content_params.get('col_width') is None:
            content_params['col_width'] = None
        self.content_params = content_params
        self.context = context
        
        # Determine height and width ---
        height = 0
        width = 0
        for field, content in self.structure.items():
            fh = field_params.get('height', 1)
            fw = field_params.get('width', 1)
            cw = content_params.get('width', 1)
            w = fw + cw + self.vspace
            if w > width:
                width = w
            vals = content['content']
            if not isinstance(vals, list):
                vals = [vals]
            ch = len(vals) * content_params.get('height', 1)
            if fh > ch:
                height += fh
            else:
                height += ch
        self.height = height
        self.width = width
    
    def load_config(self, path = None):
        """Loads config from a config.yaml file
    
        Args:
            path (str): path to a config file; may be None, then Collector object's default is used
            
        Returns:
            OrderedDict: config parsed to a dictionary
        """        
        if path is None:
            path = self.config_path
    
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

    def _merge_styles(self, style, additional_style):
        """Add and/or change styling dict
        
        Args:
            style (dict): original style dictionary
            additional_style (dict): dict with additional styling rules
        
        Returns:
            dict: merge styling dictionary
        """
        merged_style = style.copy()
        for key, value in additional_style.items():
            merged_style[key] = value
        return merged_style
    
    def process_value(self, x):
        """Evaluate string agains a context
        """
        if isinstance(x, str) and re.match('^@eval@', x):
            x = re.sub('^@eval@', '', x)
            return eval(x, None, self.context)
        else:
            return x

    def draw(self, x, y, ws, wb):
        """Draw Dictionary in a worksheet
        
        Args:
            x (int): x-coordinate (rows)
            y (int): y-coordinate (columns)
            ws (xlsxwriter.worksheet.Worksheet): worksheet to draw in
            wb (xlsxwriter.workbook.Workbook): workbook to draw in
        """
        y0 = y
        for field, data in self.structure.items():
            field_params = self._merge_styles(self.field_params, data.get('field_params', {}))
            content_params = self._merge_styles(self.content_params, data.get('content_params', {}))
            field_value = self.process_value(field)
            vspace = data.get('vspace', self.vspace)
            Field = HeaderElement(field_value, **field_params)
            Field.draw(x, y, ws, wb)
            content = data['content']
            if not isinstance(content, list):
                content = [content]
            elif content is None:
                content = ['']
            for value in content:
                value = self.process_value(value)
                Content = HeaderElement(value, **content_params)
                Content.draw(x, y  + Field.width + self.hspace, ws, wb)
                x += Content.height
            y = y0
            x += vspace

###############################################################################
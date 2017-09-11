"""The main drawing controller class"""

# Imported modules ------------------------------------------------------------

# Full imports ---
import xlsxwriter
import re

# Partial imports ---
from collections import OrderedDict
from xlsxwriter.utility import (
    xl_rowcol_to_cell,
    xl_cell_to_rowcol
)
from xlsxpandas.__internals__ import (
    validate_param
)

###############################################################################
    
class Drawer(object):
    """Elements drawer that works like a plotter
    
    This is an implementation of the drawing class.
    Drawer object is used for drawing actual drawing element in a .xlsx report.
    Its mechanics are quite simple: Drawer has its position in a matrix rows (x) * columns (y) coordinate system
    and it can be fed with drawing elements which it in turn draws (according to their attributes)
    in a place it is currently located in.
    
    Parameters
    ----------
        ws : xlsxwriter.worksheet.Worksheet
            an excel worksheet to place the drawer in
        wb : xlsxwriter.workbook.Workbook
            an excel workbook to operate on
        x : nonnegative int
            initial x-coordinate (rows) for the drawer
        y : nonnegative int
            initial y-coordinate (columns) for the drawer
        na_rep : str
            string representation of missing values (anything pandas-null); default to empty string
    
    Returns
    -------
    **Attributes**:
        * x : current x-coordinate (rows)
        * y : current y-coordinate (columns)
        * ws : worksheet to draw on
        * wb : excel workbook to operate on
        * prev_x : list of previous x-coordinates (from the most recent to the last)
        * prev_y : list of previous y-coordinates (from the most recent to the last)
        * checkpoints : OrderedDict with a set of named checkpoints
        * na_rep : string representation of missing values (anything that is pandas-null)
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def x(self):
        return self._x
    @x.setter
    def x(self, value):
        self._x = validate_param(value, 'x', int, True, 'x >= 0')
    
    @property
    def y(self):
        return self._y
    @y.setter
    def y(self, value):
        self._y = validate_param(value, 'y', int, True, 'x >= 0')
    
    @property
    def ws(self):
        return self._ws
    @ws.setter
    def ws(self, value):
        self._ws = validate_param(value, 'ws', xlsxwriter.worksheet.Worksheet)
    
    @property
    def wb(self):
        return self._wb
    @wb.setter
    def wb(self, value):
        self._wb = validate_param(value, 'wb', xlsxwriter.workbook.Workbook)
    
    @property
    def na_rep(self):
        return self._na_rep
    @na_rep.setter
    def na_rep(self, value):
        self._na_rep = validate_param(value, 'na_rep', str)
    
    # -------------------------------------------------------------------------
    
    def __init__(self, ws, wb, x = 0, y = 0, na_rep = ''):
        """Initilization method
        """
        self.x = x
        self.y = y
        self.ws = ws
        self.wb = wb
        self.checkpoints = OrderedDict()
        self.prev_x = []
        self.prev_y = []
        self.na_rep = na_rep
    
    def draw(self, elem, **kwargs):
        """Draw an element in a worksheet
        
        Parameters
        ----------
            elem : any object with a proper .draw() method 
                an object to draw on the worksheet
            **kwargs: keyword arguments passed to the invoked draw method
        """
        elem.draw(self.x, self.y, self.ws, self.wb, self.na_rep, **kwargs)
    
    def move(self, x = 0, y = 0):
        """Move drawer
        
        Move drawer by specifed number of cells horizontally and/or vertically
        
        Parameters
        ----------
            x : int
                shift value for the x-coordinate (rows)
            y : int
                shift value for the y-coordinate (columns)
        """
        self.prev_x.append(self.x)
        self.prev_y.append(self.y)
        self.x += x
        self.y += y
    
    def width(self, n = 0):
        """Get width of a previously drawn element
        
        Parameters
        ----------
            n : int
                number of steps backwards from the most recent element
        """
        try:
            return abs(self.prev_y[n] - self.prev_y[n+1])
        except IndexError:
            raise IndexError('%dth element does not yet exist.' % (n+1))
    
    def height(self, n = 0):
        """Get height of a previously drawn element
        
        Parameters
        ----------
            n : int
                number of steps backwards from the most recent element
        """
        try:
            return abs(self.prev_x[n] - self.prev_y[n+1])
        except IndexError:
            raise IndexError('%dth element does not yet exist.' % (n+1))
    
    def move_horizontal(self, back = False):
        """Move drawer horizontally
        
        This method mover drawer horizontally for the width of the last drawn object.
        If `back` is set to `True`, then it moves backward for the width of the second last drawn object.
        
        Parameters
        ----------
            back : bool
                whether to move forward or backward
        """
        if back:
            self.move(0, self.width(1))
        else:
            self.move(0, self.width(0))
    
    def move_vertical(self, back = False):
        """Move drawer vertically
        
         This method mover drawer vertically for the height of the last drawn object.
        If `back` is set to `True`, then it moves backward for the height of the second last drawn object.
        
        Parameters
        ----------
            back : bool
                whether to move forward or backward
        """
        if back:
            self.move(self.height(1), 0)
        else:
            self.move(self.height(0), 0)
    
    def add_checkpoint(self, name):
        """Adds current position as a named checkpoint
        
        Parameters
        ----------
            name : str
                name for the checkpoint
        """
        self.checkpoints[name] = (self.x, self.y)
    
    def reset(self, x = 0, y = 0, checkpoint = None):
        """Reset Drawer position
        
        If checkpoint name is provided, then the Drawer is reset to the checkpoint.
        Otherwise it is reset to the given x and y coordinates.
        If `x` or `y` is `None` then this dimension is not changed.
        
        Parameters
        ----------
            name : str or None
                name of a checkpoint to fallback to
            x : int or None
                new x-coordinate to assign; no change if `None`
            y : int or None
                new y-coordinate to assign; no change if `None`
        """
        self.prev_x.append(self.x)
        self.prev_y.append(self.y)
        
        if checkpoint is not None:
            if x is not None:
                self.x = self.checkpoints[checkpoint][0]
            if y is not None:
                self.y = self.checkpoints[checkpoint][1]
        else:
            if x is not None:
                self.x = x
            if y is not None:
                self.y = y
        
    def fallback(self, n):
        """Fall back to the nth previous step
        
        Parameters
        ----------
            n : int
                number of steps to fall back. Negative values iterate from the historically first position.
        """
        self.reset(x = self.prev_x[-n], y = self.prev_y[-n])
    
    def xl_position(self, x = 0, y = 0):
        """Get Drawer's position in the excel notation
        
        Parameters
        ----------
            x : int
                number of rows to shift when determining the position
            y : int
                number of columns to shift when determinin the position
        """
        return xl_rowcol_to_cell(self.x + x, self.y + y)
    
    def xl_column(self, y = 0):
        """Get Drawer's current column in the excel notation
        
        Parameters
        ----------
            y : int
                number of columns to shift when determining the position
        """
        return re.sub('[0-9]', '', self.xl_position(y = y))
    
    def xl_row(self, x = 0):
        """Get Drawer's current row in the excel notation
        
        Parameters
        ----------
            x : int
                number of rows to shift when determining position
        """
        return re.sub('[A-Z]', '', self.xl_position(x = x))
    
    @staticmethod
    def xl2coords(rng):
        """Translate an excel range string to matrix coordinates
        
        Parameters
        ----------
            rng : str
                an excel range string
        
        Returns
        -------
           coords : length 2 tuple of ints
        """
        return xl_cell_to_rowcol(rng)
    
    def xl_set(self, rng):
        """Set the drawer's position using an excel range string
        
        Parameters
        ----------
            rng : str
                excel position to set to
        """
        self.x, self.y = self.xl2coords(rng)       
    
###############################################################################
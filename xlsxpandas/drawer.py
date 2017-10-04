"""The main drawing controller class"""

# Imported modules ------------------------------------------------------------

# Full imports ---
import xlsxwriter
import re

# Partial imports ---
from collections import OrderedDict, deque
from xlsxwriter.utility import (
    xl_rowcol_to_cell,
    xl_cell_to_rowcol
)
from xlsxpandas.__internals__ import (
    validate_param
)

###############################################################################


class Drawer(object):
    """Elements drawer that works like a plotter.

    This is an implementation of the drawing class.
    Drawer object is used for drawing actual drawing element in a .xlsx report.
    Its mechanics are quite simple: Drawer has its position
    in a matrix rows (x) * columns (y) coordinate system
    and it can be fed with drawing elements which it in turn draws
    (according to their attributes) in a place it is currently located in.
    """

    # -------------------------------------------------------------------------

    @property
    def x(self):
        """nonnegative int: Current x-coordinate (rows) of the drawer."""
        return self._x

    @property
    def y(self):
        """nonegative int: Current y-coordiate (columns) of the drawer."""
        return self._y

    @property
    def ws(self):
        """xlsxwriter.worksheet.Worksheet: Excel worksheet the drawer is in."""
        return self._ws

    @ws.setter
    def ws(self, value):
        self._ws = validate_param(value, 'ws', xlsxwriter.worksheet.Worksheet)

    @property
    def wb(self):
        """xlsxwriter.workbook.Workbook: Excel workbook the drawer is in."""
        return self._wb

    @wb.setter
    def wb(self, value):
        self._wb = validate_param(value, 'wb', xlsxwriter.workbook.Workbook)

    @property
    def na_rep(self):
        """str: String representation of missing value in the worksheet."""
        return self._na_rep

    @na_rep.setter
    def na_rep(self, value):
        self._na_rep = validate_param(value, 'na_rep', str)

    @property
    def widths(self):
        """list: List of widths of drawn objects."""
        return self._widths

    @property
    def heights(self):
        """list: List of heights of drawn object."""
        return self._heights

    # -------------------------------------------------------------------------

    def __init__(self, ws, wb, x=0, y=0, na_rep='', memlen=10):
        """Initilization method.

        Parameters
        ----------
        ws : xlsxwriter.worksheet.Worksheet
            Excel worksheet to place the drawer in.
        wb : xlsxwriter.workbook.Workbook
            Excel workbook to operate on.
        x : nonnegative int
            Initial x-coordinate (rows) for the drawer.
        y : nonnegative int
            Initial y-coordinate (columns) for the drawer.
        na_rep : str
            String representation of missing values (anything pandas-null).
            Defaults to empty string.
        memlen : int
            Maximum length of stored previous widths and heights
            of drawn objects.
        """
        self._x = x
        self._y = y
        self.ws = ws
        self.wb = wb
        self.checkpoints = OrderedDict()
        self.na_rep = na_rep
        self._widths = deque([], maxlen=memlen)
        self._heights = deque([], maxlen=memlen)

    def draw(self, elem, **kwargs):
        """Draw an element in a worksheet.

        Parameters
        ----------
            elem : any object with a proper `draw` method
            and `width` and `height` properties.
                An object to draw on the worksheet.
            **kwargs
                Keyword arguments passed to the invoked draw method.
        """
        elem.draw(self.x, self.y, self.ws, self.wb, self.na_rep, **kwargs)
        self.widths.append(elem.width)
        self.heights.append(elem.height)

    def move(self, x=0, y=0):
        """Move drawer.

        Move drawer by specifed number of cells horizontally and/or vertically.

        Parameters
        ----------
            x : int
                Shift value for the x-coordinate (rows).
            y : int
                Shift value for the y-coordinate (columns).
        """
        self._x += x
        self._y += y

    def width(self, n=0):
        """Get width of a previously drawn element.

        Parameters
        ----------
            n : int
                Number of steps backwards from the most recent element.
        """
        try:
            return self.widths[-(1 + n)]
        except IndexError:
            raise IndexError('%dth element does not yet exist.' % (n+1))

    def height(self, n=0):
        """Get height of a previously drawn element.

        Parameters
        ----------
            n : int
                Number of steps backwards from the most recent element.
        """
        try:
            return self.heights[-(1 + n)]
        except IndexError:
            raise IndexError('%dth element does not yet exist.' % (n+1))

    def move_horizontal(self, back=False):
        """Move drawer horizontally.

        This method mover drawer horizontally for the width
        of the last drawn object.  If `back` is set to `True`,
        then it moves backward for the width of the second last drawn object.

        Parameters
        ----------
            back : bool
                Whether to move forward or backward.
        """
        if back:
            self.move(0, self.width(1))
        else:
            self.move(0, self.width(0))

    def move_vertical(self, back=False):
        """Move drawer vertically.

        This method mover drawer vertically for the height
        of the last drawn object. If `back` is set to `True`,
        then it moves backward for the height of the second last drawn object.

        Parameters
        ----------
            back : bool
                Whether to move forward or backward.
        """
        if back:
            self.move(self.height(1), 0)
        else:
            self.move(self.height(0), 0)

    def add_checkpoint(self, name):
        """Adds current position as a named checkpoint.

        Parameters
        ----------
            name : str
                Name for the checkpoint.
        """
        self.checkpoints[name] = (self.x, self.y)

    def reset(self, x=0, y=0, checkpoint=None):
        """Reset Drawer position.

        If checkpoint name is provided,
        then the Drawer is reset to the checkpoint.
        Otherwise it is reset to the given x and y coordinates.
        If `x` or `y` is `None` then this dimension is not changed.

        Parameters
        ----------
            name : str or None
                Name of a checkpoint to fallback to.
            x : int or None
                New x-coordinate to assign; no change if `None`.
            y : int or None
                New y-coordinate to assign; no change if `None`.
        """
        if checkpoint is not None:
            if x is not None:
                self._x = self.checkpoints[checkpoint][0]
            if y is not None:
                self._y = self.checkpoints[checkpoint][1]
        else:
            if x is not None:
                self._x = x
            if y is not None:
                self._y = y

    def xl_position(self, x=0, y=0):
        """Get Drawer's position in the excel notation.

        Parameters
        ----------
            x : int
                Number of rows to shift when determining the position.
            y : int
                Number of columns to shift when determinin the position.
        """
        return xl_rowcol_to_cell(self.x + x, self.y + y)

    def xl_column(self, y=0):
        """Get Drawer's current column in the excel notation.

        Parameters
        ----------
            y : int
                Number of columns to shift when determining the position.
        """
        return re.sub('[0-9]', '', self.xl_position(y=y))

    def xl_row(self, x=0):
        """Get Drawer's current row in the excel notation

        Parameters
        ----------
            x : int
                Number of rows to shift when determining position.
        """
        return re.sub('[A-Z]', '', self.xl_position(x=x))

    def xl_upleft(self, x=0, y=0):
        """Get upper left corner coordinates of the last drawn object.

        Parameters
        ----------
            x : int
                Number of rows to shift when determining the position.
            y : int
                Number of columns to shift when determinin the position.
        """
        return self.xl_position(x=x, y=y)

    def xl_loright(self, x=0, y=0):
        """Get lower right corner coordinates of the last drawn object.

        Parameters
        ----------
            x : int
                Number of rows to shift when determining the position.
            y : int
                Number of columns to shift when determinin the position.
        """
        return self.xl_position(x=self.height()-1+x, y=self.width()-1+y)

    def xl_range(self):
        """Get range covered by the last drawn element.
        """
        return self.xl_upleft() + ':' + self.xl_loright()

    @staticmethod
    def xl2coords(rng):
        """Translate an excel range string to matrix coordinates.

        Parameters
        ----------
            rng : str
                An excel range string.

        Returns
        -------
           tuple
               length 2 tuple of ints with appropriate coordinates.
        """
        return xl_cell_to_rowcol(rng)

    def xl_set(self, rng):
        """Set the drawer's position using an excel range string.

        Parameters
        ----------
            rng : str
                Excel position to set to.
        """
        self.x, self.y = self.xl2coords(rng)

###############################################################################

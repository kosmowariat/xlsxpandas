"""Internal functions used throught the package"""

# Imported moduels -----------------------------------------------------------


###############################################################################

def validate_param(x, name, expected_type, coerce = False, *args):
    """Validate a parameter
    
    Parameters
    ----------
        x : anything
            value to be checked
        name : str
            name of the parameter
        expected_type: type or a tuple of types
            type(s) to check for
        coerce : bool or function
            should attempt to coerce `x` first; arbitrary coercion function may be passed
        *args : str
            strings with arbitrary expressions for validating the parameter
    """
    # Optional type coercion
    if coerce is True:
        coerce = expected_type
    if coerce:
        try:
            x = coerce(x)
        except:
            raise ValueError('`%s` can not be coerced to %r.' % (name, expected_type))
    
    # Test and return
    assert isinstance(x, expected_type), '`%s` is not %s' % (name, str(expected_type))
    unmet_conditions = []
    if args:
        for expr in args:
            condition = eval(expr)
            if not isinstance(condition, bool):
                raise ValueError('`%s` does not evaluate to %r' % (expr, bool))
            if not condition:
                unmet_conditions.append(expr)
    if unmet_conditions:
        msg = '`%s` does not satisfy the conditions:\n' % name
        msg += '\n'.join(unmet_conditions)
        raise AssertionError(msg)
    return x
        
#This is merely an error handling code for gspread sheets. It came from SnoopJ on LiberaChat's IRC.

import tester.py
from contextlib import contextmanager
from pprint import pprint

from gspread.exceptions import APIError


@contextmanager
def better_errors():
    try:
        yield
    except APIError as exc:
        response = exc.response
        try:
            pprint(response.json())
        except:
            print(response)
        raise exc


gray = cellFormat(backgroundColor=color(153, 153, 153))

# any APIErrors caused by this block will be reported more usefully
with better_errors():
    test()

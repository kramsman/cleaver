""" new version which accepts a path as input and uses pathlib not os"""

import pathlib
import pymsgbox
import inspect


def calling_func(level=0):
    """ returns the various levels of calling function.  0 is current, 1 is caller of current, etc """
    try:
        func = f"'{inspect.stack()[level][3]}', line #: {inspect.stack()[level][2]}"
    except Exception:
        func = f"** error ** inspect level too deep: {str(level)} called from {inspect.stack()[level][3]}"
    return func


def exit_yes(msg: str, title: str = None, *, errmsg: str = None) -> None:
    """ exits program after giving user a popup window and raising an error. """
    msg = (msg + "\n\n\nExiting." +
           f"\n\nCalled from {calling_func(level=3)}"
           f"\nCalled from {calling_func(level=2)}"
           f"\nCalled from {calling_func(level=1)}"
           )
    if not errmsg:
        errmsg = msg.replace("\n", " ")  # dont fill the console with linefeeds
    if not title:
        title = "** Exiting Program **"
    pymsgbox.alert(msg, title)
    raise Exception(errmsg)


def bad_path_exit(my_path, msg=None):
    """ checks for directory existence and exits if not found"""
    if msg is None:
        msg = f"Directory:\n\n'{my_path}'\n\ndoes not exist."
    if isinstance(my_path, str):
        my_path = pathlib.Path(my_path)
    my_path = my_path.expanduser()
    # if not Path(os.path.expanduser(path)).exists():  # need expanduser for ~; only os works (not pathlib)
    if not my_path.exists():  # need expanduser for ~; only os works (not pathlib)
        # pymsgbox.alert(msg, "** Exiting via bad_path_exit **")
        # # FIXME: close TKINTER window here.  https://stackoverflow.com/questions/8009176/function-to-close-the-window-in-tkinter
        # exit()
        exit_yes(msg)


if __name__ == "__main__":
    bad_path_exit('/Users/Denise/Downloads/all-users-2023-11-10.csv')
    bad_path_exit('~/Downloads/all-users-2023-11-10.csv')
    p = pathlib.Path('/Users/Denise/Downloads/all-users-2023-11-10.csv')
    bad_path_exit(p)
    bad_path_exit('/xxx/all-users-2023-11-10.csv')
    a=1

# wdlg

This package is a small python library for windows.
It provides functions for creating file/directory selection dialogs and message boxes.

This package is written entirely in Python.
It uses `ctypes` module to display dialogs and has no dependencies on `tkinter` or outside the standard library.

Function and argument names are almost the same as those of `tkinter`.
I feel that it's easy to port code from `tkinter` to `wdlg`, and vice versa.

## Installation

```
pip install wdlg
```

## Example

```python
# it's easy to port code from tkinter to wdlg.

# from tkinter import filedialog
# from tkinter import messagebox
from wdlg import filedialog
from wdlg import messagebox

dir_path = filedialog.askdirectory(initialdir="C:\\Users")
messagebox.showinfo(title="Hello World", message=dir_path)
```

## License

BSD Zero Clause License

see LICENSE.txt.

"""for creating the message boxes.

this module provides functions for creating message boxes.
"""
__all__ = [
    "showinfo",
    "showwarning",
    "showerror",
    "askokcancel",
    "askquestion",
    "askretrycancel",
    "askyesno",
    "askyesnocancel",
]
from ctypes import POINTER
from ctypes import byref
from ctypes import c_int32
from ctypes import c_void_p
from ctypes import c_wchar_p
from ctypes import oledll

oledll.comctl32.TaskDialog.argtypes = [
    c_void_p,
    c_void_p,
    c_wchar_p,
    c_wchar_p,
    c_wchar_p,
    c_int32,
    c_wchar_p,
    POINTER(c_int32),
]


def taskdialog(
    hwndOwner: c_void_p = None,
    hInstance: c_void_p = None,
    pszWindowTitle: str = None,
    pszMainInstruction: str = None,
    pszContent: str = None,
    dwCommonButtons: int = 0x1,
    pszIcon: c_wchar_p = None,
) -> int:
    """creates, displays, and operates a task dialog.

    Args:
        hwndOwner:
            Handle to the owner window of the task dialog to be created.
            If this parameter is None, the task dialog has no owner window.

        hInstance:
            Handle to the module that contains
            the icon resource identified by the pszIcon member.
            If this parameter is None,
            pszIcon must be None or a system resource identifier.

        pszWindowTitle:
            The task dialog title.
            If this parameter is None, the filename of the executable program is used.

        pszMainInstruction:
            The main instruction.
            This parameter can be None if no main instruction is wanted.

        pszContent:
            The additional text.
            Can be None if no additional text is wanted.

        dwCommonButtons:
            Specifies the push buttons displayed in the dialog box.
            This parameter may be a combination of flags from the following group.::

                0x01    TDCBF_OK_BUTTON (return IDOK = 0x01).
                0x02    TDCBF_YES_BUTTON (return IDYES = 0x06).
                0x04    TDCBF_NO_BUTTON (return IDNO = 0x07).
                0x08    TDCBF_CANCEL_BUTTON (return IDCANCEL = 0x02).
                0x10    TDCBF_RETRY_BUTTON (return IDRETRY = 0x04).
                0x20    TDCBF_CLOSE_BUTTON (return IDCLOSE = 0x08).

        pszIcon:
            Pointer to a string that identifies the icon to display in the dialog.
            It can take following values.::

                None                 no icon.
                c_wchar_p(0xFFFF)    WARNING_ICON (⚠️).
                c_wchar_p(0xFFFE)    ERROR_ICON (❌).
                c_wchar_p(0xFFFD)    INFORMATION_ICON (ℹ️).
                c_wchar_p(0xFFFC)    SHIELD_ICON (🛡️).

    Returns:
        int: return one of the following values::

            0x00    Function call failed.
            0x01    OK button was selected (IDOK).
            0x02    Cancel button was selected,
                    or the user clicked on the close window button (IDCANCEL).
            0x04    Retry button was selected (IDRETRY).
            0x06    Yes button was selected (IDYES).
            0x07    No button was selected (IDNO).
            0x08    Close button was selected (IDCLOSE).

    Raises:
        OSError

    https://docs.microsoft.com/en-us/windows/win32/api/commctrl/nf-commctrl-taskdialog
    """
    pnButton = c_int32(0)
    oledll.comctl32.TaskDialog(
        hwndOwner,
        hInstance,
        pszWindowTitle,
        pszMainInstruction,
        pszContent,
        dwCommonButtons,
        pszIcon,
        byref(pnButton),
    )
    return pnButton.value


def showinfo(
    hwnd: c_void_p = None,
    title: str = None,
    main_msg: str = None,
    message: str = None,
) -> None:
    """Show an info message.

    Args:
        hwnd: A handle to the window that owns the dialog box.
        title: A string to be used for window title.
        main_msg: A string to be used for the main instruction.
        message: A string used for additional text.
    """
    taskdialog(hwnd, None, title, main_msg, message, 0x01, c_wchar_p(0xFFFD))
    return None


def showwarning(
    hwnd: c_void_p = None,
    title: str = None,
    main_msg: str = None,
    message: str = None,
) -> None:
    """Show a warning message.

    Args:
        hwnd: A handle to the window that owns the dialog box.
        title: A string to be used for window title.
        main_msg: A string to be used for the main instruction.
        message: A string used for additional text.
    """
    taskdialog(hwnd, None, title, main_msg, message, 0x01, c_wchar_p(0xFFFF))
    return None


def showerror(
    hwnd: c_void_p = None,
    title: str = None,
    main_msg: str = None,
    message: str = None,
) -> None:
    """Show a error messsage.

    Args:
        hwnd: A handle to the window that owns the dialog box.
        title: A string to be used for window title.
        main_msg: A string to be used for the main instruction.
        message: A string used for additional text.
    """
    taskdialog(hwnd, None, title, main_msg, message, 0x01, c_wchar_p(0xFFFE))
    return None


def askokcancel(
    hwnd: c_void_p = None,
    title: str = None,
    main_msg: str = None,
    message: str = None,
) -> bool:
    """Ask if operation should proceed, return true if the answer is ok.

    Args:
        hwnd: A handle to the window that owns the dialog box.
        title: A string to be used for window title.
        main_msg: A string to be used for the main instruction.
        message: A string used for additional text.

    Returns:
        Return true, if the answer is ok.
    """
    # IDI_QUESTION = 32514
    # see the docs of LoadIconW function (Win32 API)
    ret = taskdialog(hwnd, None, title, main_msg, message, 0x09, c_wchar_p(32514))
    if ret == 0x01:
        return True
    else:
        return False


def askquestion(
    hwnd: c_void_p = None,
    title: str = None,
    main_msg: str = None,
    message: str = None,
) -> bool:
    """Ask a question, return true if the answer is yes.

    Args:
        hwnd: A handle to the window that owns the dialog box.
        title: A string to be used for window title.
        main_msg: A string to be used for the main instruction.
        message: A string used for additional text.

    Returns:
        Return true, if the answer is yes.
    """
    ret = taskdialog(hwnd, None, title, main_msg, message, 0x06, c_wchar_p(32514))
    if ret == 0x06:
        return True
    else:
        return False


def askretrycancel(
    hwnd: c_void_p = None,
    title: str = None,
    main_msg: str = None,
    message: str = None,
) -> bool:
    """Ask if operation should be retried, return true if the answer is yes.

    Args:
        hwnd: A handle to the window that owns the dialog box.
        title: A string to be used for window title.
        main_msg: A string to be used for the main instruction.
        message: A string used for additional text.

    Returns:
        Return true, if the answer is yes.
    """
    ret = taskdialog(hwnd, None, title, main_msg, message, 0x18, c_wchar_p(0xFFFF))
    if ret == 0x04:
        return True
    else:
        return False


def askyesno(
    hwnd: c_void_p = None,
    title: str = None,
    main_msg: str = None,
    message: str = None,
) -> bool:
    """Ask a question, return true if the answer is yes.

    Args:
        hwnd: A handle to the window that owns the dialog box.
        title: A string to be used for window title.
        main_msg: A string to be used for the main instruction.
        message: A string used for additional text.

    Returns:
        Return true, if the answer is yes.
    """
    return askquestion(title, main_msg, message, hwnd)


def askyesnocancel(
    hwnd: c_void_p = None,
    title: str = None,
    main_msg: str = None,
    message: str = None,
) -> bool:
    """Ask a question, return true if the answer is yes, return None if cancelled.

    Args:
        hwnd: A handle to the window that owns the dialog box.
        title: A string to be used for window title.
        main_msg: A string to be used for the main instruction.
        message: A string used for additional text.

    Returns:
        Return true if the answer is yes.
        Return false if the answer is no.
        Return None if the answer is cancel.
    """
    ret = taskdialog(hwnd, None, title, main_msg, message, 0x0E, c_wchar_p(0xFFFF))
    if ret == 0x06:
        return True
    elif ret == 0x07:
        return False
    else:
        return None

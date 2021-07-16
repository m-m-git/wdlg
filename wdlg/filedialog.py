"""for creating file/directory dialogs.

This module provides functions for creating file/directory selection dialogs.
"""
__all__ = ["askdirectory", "askopenfilename", "askopenfilenames", "asksaveasfilename"]
from contextlib import contextmanager
from ctypes import HRESULT
from ctypes import POINTER
from ctypes import WINFUNCTYPE
from ctypes import Structure
from ctypes import byref
from ctypes import c_uint8
from ctypes import c_uint16
from ctypes import c_uint32
from ctypes import c_void_p
from ctypes import c_wchar_p
from ctypes import oledll
from typing import Generator
from typing import List
from typing import Tuple

# COM initialization flags; passed to CoInitialize.
COINIT_APARTMENTTHREADED = 0x2
COINIT_DISABLE_OLE1DDE = 0x4

# flags to indicate the execution contexts; passed to CoCreateInstance.
CLSCTX_INPROC_SERVER = 0x1

# IFileOpenDialog options; passed to IFileOpenDialog::SetOptions.
FOS_OVERWRITEPROMPT = 0x2  # default save
FOS_PICKFOLDERS = 0x20
FOS_ALLOWMULTISELECT = 0x200
FOS_PATHMUSTEXIST = 0x800  # default
FOS_FILEMUSTEXIST = 0x1000  # default open

# flags to indicate the form of an display name; passed to IShellItem::GetDisplayName.
SIGDN_FILESYSPATH = 0x80058000

_wf = WINFUNCTYPE(HRESULT)


class GUID(Structure):
    _fields_ = [
        ("Data1", c_uint32),
        ("Data2", c_uint16),
        ("Data3", c_uint16),
        ("Data4", c_uint8 * 8),
    ]

    @staticmethod
    def from_str(str_guid: str) -> "GUID":
        ret = GUID()
        oledll.ole32.CLSIDFromString(c_wchar_p(str_guid), byref(ret))
        return ret


class COMDLG_FILTERSPEC(Structure):
    _fields_ = [
        ("pszName", c_wchar_p),
        ("pszSpec", c_wchar_p),
    ]

    @staticmethod
    def from_tuple(filetype: Tuple[str, str]) -> "COMDLG_FILTERSPEC":
        return COMDLG_FILTERSPEC(filetype[0], filetype[1])


class IFileOpenDialog(Structure):
    pass


class IFileOpenDialogVtbl(Structure):
    pass


class IFileSaveDialog(Structure):
    pass


class IFileSaveDialogVtbl(Structure):
    pass


class IShellItem(Structure):
    pass


class IShellItemVtbl(Structure):
    pass


class IShellItemArray(Structure):
    pass


class IShellItemArrayVtbl(Structure):
    pass


IShellItem._fields_ = [("lpVtbl", POINTER(IShellItemVtbl))]
IShellItemVtbl._fields_ = [
    ("QueryInterface", _wf),
    ("AddRef", WINFUNCTYPE(c_uint32, POINTER(IShellItem))),
    ("Release", WINFUNCTYPE(c_uint32, POINTER(IShellItem))),
    ("BindToHandler", _wf),
    ("GetParent", _wf),
    (
        "GetDisplayName",
        WINFUNCTYPE(HRESULT, POINTER(IShellItem), c_uint32, POINTER(c_wchar_p)),
    ),
    ("GetAttributes", _wf),
    ("Compare", _wf),
]

IShellItemArray._fields_ = [("lpVtbl", POINTER(IShellItemArrayVtbl))]
IShellItemArrayVtbl._fields_ = [
    ("QueryInterface", _wf),
    ("AddRef", WINFUNCTYPE(c_uint32, POINTER(IShellItemArray))),
    ("Release", WINFUNCTYPE(c_uint32, POINTER(IShellItemArray))),
    ("BindToHandler", _wf),
    ("GetPropertyStore", _wf),
    ("GetPropertyDescriptionList", _wf),
    ("GetAttributes", _wf),
    ("GetCount", WINFUNCTYPE(HRESULT, POINTER(IShellItemArray), POINTER(c_uint32))),
    (
        "GetItemAt",
        WINFUNCTYPE(
            HRESULT, POINTER(IShellItemArray), c_uint32, POINTER(POINTER(IShellItem))
        ),
    ),
    ("EnumItems", _wf),
]

IFileOpenDialog._fields_ = [("lpVtbl", POINTER(IFileOpenDialogVtbl))]
IFileOpenDialogVtbl._fields_ = [
    ("QueryInterface", _wf),
    ("AddRef", WINFUNCTYPE(c_uint32, POINTER(IFileOpenDialog))),
    ("Release", WINFUNCTYPE(c_uint32, POINTER(IFileOpenDialog))),
    ("Show", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), c_void_p)),
    (
        "SetFileTypes",
        WINFUNCTYPE(
            HRESULT, POINTER(IFileOpenDialog), c_uint32, POINTER(COMDLG_FILTERSPEC)
        ),
    ),
    ("SetFileTypeIndex", _wf),
    ("GetFileTypeIndex", _wf),
    ("Advise", _wf),
    ("Unadvise", _wf),
    ("SetOptions", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), c_uint32)),
    ("GetOptions", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(c_uint32))),
    ("SetDefaultFolder", _wf),
    ("SetFolder", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(IShellItem))),
    ("GetFolder", _wf),
    ("GetCurrentSelection", _wf),
    ("SetFileName", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), c_wchar_p)),
    ("GetFileName", _wf),
    ("SetTitle", _wf),
    ("SetOkButtonLabel", _wf),
    ("SetFileNameLabel", _wf),
    (
        "GetResult",
        WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(POINTER(IShellItem))),
    ),
    ("AddPlace", _wf),
    ("SetDefaultExtension", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), c_wchar_p)),
    ("Close", _wf),
    ("SetClientGuid", _wf),
    ("ClearClientData", _wf),
    ("SetFilter", _wf),
    (
        "GetResults",
        WINFUNCTYPE(
            HRESULT, POINTER(IFileOpenDialog), POINTER(POINTER(IShellItemArray))
        ),
    ),
    ("GetSelectedItems", _wf),
]

IFileSaveDialog._fields_ = [("lpVtbl", POINTER(IFileSaveDialogVtbl))]
IFileSaveDialogVtbl._fields_ = [
    ("QueryInterface", _wf),
    ("AddRef", WINFUNCTYPE(c_uint32, POINTER(IFileSaveDialog))),
    ("Release", WINFUNCTYPE(c_uint32, POINTER(IFileSaveDialog))),
    ("Show", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), c_void_p)),
    (
        "SetFileTypes",
        WINFUNCTYPE(
            HRESULT, POINTER(IFileSaveDialog), c_uint32, POINTER(COMDLG_FILTERSPEC)
        ),
    ),
    ("SetFileTypeIndex", _wf),
    ("GetFileTypeIndex", _wf),
    ("Advise", _wf),
    ("Unadvise", _wf),
    ("SetOptions", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), c_uint32)),
    ("GetOptions", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(c_uint32))),
    ("SetDefaultFolder", _wf),
    ("SetFolder", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(IShellItem))),
    ("GetFolder", _wf),
    ("GetCurrentSelection", _wf),
    ("SetFileName", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), c_wchar_p)),
    ("GetFileName", _wf),
    ("SetTitle", _wf),
    ("SetOkButtonLabel", _wf),
    ("SetFileNameLabel", _wf),
    (
        "GetResult",
        WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(POINTER(IShellItem))),
    ),
    ("AddPlace", _wf),
    ("SetDefaultExtension", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), c_wchar_p)),
    ("Close", _wf),
    ("SetClientGuid", _wf),
    ("ClearClientData", _wf),
    ("SetFilter", _wf),
    ("SetSaveAsItem", _wf),
    ("SetProperties", _wf),
    ("SetCollectedProperties", _wf),
    ("GetProperties", _wf),
    ("ApplyProperties", _wf),
]

CLSID_FileOpenDialog = GUID.from_str("{DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7}")
CLSID_FileSaveDialog = GUID.from_str("{C0B4E2F3-BA21-4773-8DBA-335EC946EB8B}")
CLSID_ShellItem = GUID.from_str("{9ac9fbe1-e0a2-4ad6-b4ee-e212013ea917}")
IID_IFileOpenDialog = GUID.from_str("{d57c7288-d4ad-4768-be02-9d969532d960}")
IID_IFileSaveDialog = GUID.from_str("{84bccd23-5fde-4cdb-aea4-af64b83d78ab}")
IID_IShellItem = GUID.from_str("{43826d1e-e718-42ee-bc55-a1e261c37bfe}")

oledll.ole32.CLSIDFromString.argtypes = [c_wchar_p, POINTER(GUID)]
oledll.ole32.CoInitializeEx.argtypes = [c_void_p, c_uint32]
oledll.ole32.CoCreateInstance.argtypes = [GUID, c_void_p, c_uint32, GUID, c_void_p]
oledll.ole32.CoTaskMemFree.argtypes = [c_void_p]
oledll.shell32.SHCreateItemFromParsingName.argtypes = [
    c_wchar_p,
    c_void_p,
    GUID,
    c_void_p,
]


@contextmanager
def _COM() -> Generator[None, None, None]:
    oledll.ole32.CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE)
    try:
        yield None
    finally:
        oledll.ole32.CoUninitialize()


@contextmanager
def _get_pIFileOpenDialog() -> Generator[POINTER(IFileOpenDialog), None, None]:
    pdlg = POINTER(IFileOpenDialog)()
    oledll.ole32.CoCreateInstance(
        CLSID_FileOpenDialog,
        None,
        CLSCTX_INPROC_SERVER,
        IID_IFileOpenDialog,
        byref(pdlg),
    )
    try:
        yield pdlg
    finally:
        pdlg.contents.lpVtbl.contents.Release(pdlg)


@contextmanager
def _get_pIFileSaveDialog() -> Generator[POINTER(IFileSaveDialog), None, None]:
    pdlg = POINTER(IFileSaveDialog)()
    oledll.ole32.CoCreateInstance(
        CLSID_FileSaveDialog,
        None,
        CLSCTX_INPROC_SERVER,
        IID_IFileSaveDialog,
        byref(pdlg),
    )
    try:
        yield pdlg
    finally:
        pdlg.contents.lpVtbl.contents.Release(pdlg)


@contextmanager
def _get_pIShellItem() -> Generator[POINTER(IShellItem), None, None]:
    pitem = POINTER(IShellItem)()
    try:
        yield pitem
    finally:
        # release if not null pointer.
        if pitem:
            pitem.contents.lpVtbl.contents.Release(pitem)


@contextmanager
def _get_pIShellItemArray() -> Generator[POINTER(IShellItemArray), None, None]:
    psia = POINTER(IShellItemArray)()
    try:
        yield psia
    finally:
        if psia:
            psia.contents.lpVtbl.contents.Release(psia)


def askdirectory(hwnd: c_void_p = None, initialdir: str = None) -> str:
    """Ask for a directory, and return the path.

    Args:
        hwnd: A handle to the window that owns the dialog box.
        initialdir: The initial directory.

    Returns:
        The directory path or None if cancel button selected.
    """
    with _COM():
        with _get_pIFileOpenDialog() as pdlg:
            options = c_uint32()
            pdlg.contents.lpVtbl.contents.GetOptions(pdlg, byref(options))
            pdlg.contents.lpVtbl.contents.SetOptions(
                pdlg, options.value | FOS_PICKFOLDERS
            )

            if initialdir:
                with _get_pIShellItem() as pitem:
                    oledll.shell32.SHCreateItemFromParsingName(
                        initialdir, None, IID_IShellItem, byref(pitem)
                    )
                    pdlg.contents.lpVtbl.contents.SetFolder(pdlg, pitem)

            try:
                pdlg.contents.lpVtbl.contents.Show(pdlg, hwnd)
            except OSError:
                return None
            with _get_pIShellItem() as pitem:
                pfilepath = c_wchar_p()
                pdlg.contents.lpVtbl.contents.GetResult(pdlg, byref(pitem))
                pitem.contents.lpVtbl.contents.GetDisplayName(
                    pitem, SIGDN_FILESYSPATH, byref(pfilepath)
                )
                ret = pfilepath.value
                oledll.ole32.CoTaskMemFree(pfilepath)
    return ret


def askopenfilename(
    hwnd: c_void_p = None,
    initialdir: str = None,
    initialfile: str = None,
    filetypes: List[Tuple[str, str]] = None,
) -> str:
    """Ask for a file to open, and return the path.

    Args:
        hwnd: A handle to the window that owns the dialog box.
        initialdir: The directory that the dialog starts in.
        initialfile: The file name used to initialize the File Name edit control.
        filetypes: The file types that the dialog can open.

    Returns:
        The file path or None if cancel button selected.
    """
    with _COM():
        with _get_pIFileOpenDialog() as pdlg:
            if initialdir:
                with _get_pIShellItem() as pitem:
                    oledll.shell32.SHCreateItemFromParsingName(
                        initialdir, None, IID_IShellItem, byref(pitem)
                    )
                    pdlg.contents.lpVtbl.contents.SetFolder(pdlg, pitem)

            if initialfile:
                pdlg.contents.lpVtbl.contents.SetFileName(pdlg, initialfile)

            if filetypes:
                rgSpec = (COMDLG_FILTERSPEC * len(filetypes))()
                for i, ft in enumerate(filetypes):
                    rgSpec[i] = COMDLG_FILTERSPEC.from_tuple(ft)
                pdlg.contents.lpVtbl.contents.SetFileTypes(pdlg, len(filetypes), rgSpec)
                pdlg.contents.lpVtbl.contents.SetDefaultExtension(pdlg, "")

            try:
                pdlg.contents.lpVtbl.contents.Show(pdlg, hwnd)
            except OSError:
                return None
            with _get_pIShellItem() as pItem:
                pfilepath = c_wchar_p()
                pdlg.contents.lpVtbl.contents.GetResult(pdlg, byref(pItem))
                pItem.contents.lpVtbl.contents.GetDisplayName(
                    pItem, SIGDN_FILESYSPATH, byref(pfilepath)
                )
                ret = pfilepath.value
                oledll.ole32.CoTaskMemFree(pfilepath)
    return ret


def askopenfilenames(
    hwnd: c_void_p = None,
    initialdir: str = None,
    initialfile: str = None,
    filetypes: List[Tuple[str, str]] = None,
) -> List[str]:
    """Ask for multiple files to open, and return the list of paths.

    Args:
        hwnd: A handle to the window that owns the dialog box.
        initialdir: The directory that the dialog starts in.
        initialfile: The file name used to initialize the File Name edit control.
        filetypes: The file types that the dialog can open.

    Returns:
        The list of file paths or None if cancel button selected.
    """
    with _COM():
        with _get_pIFileOpenDialog() as pdlg:
            options = c_uint32()
            pdlg.contents.lpVtbl.contents.GetOptions(pdlg, byref(options))
            pdlg.contents.lpVtbl.contents.SetOptions(
                pdlg, options.value | FOS_ALLOWMULTISELECT
            )

            if initialdir:
                with _get_pIShellItem() as pitem:
                    oledll.shell32.SHCreateItemFromParsingName(
                        initialdir, None, IID_IShellItem, byref(pitem)
                    )
                    pdlg.contents.lpVtbl.contents.SetFolder(pdlg, pitem)

            if initialfile:
                pdlg.contents.lpVtbl.contents.SetFileName(pdlg, initialfile)

            if filetypes:
                rgSpec = (COMDLG_FILTERSPEC * len(filetypes))()
                for i, ft in enumerate(filetypes):
                    rgSpec[i] = COMDLG_FILTERSPEC.from_tuple(ft)
                pdlg.contents.lpVtbl.contents.SetFileTypes(pdlg, len(filetypes), rgSpec)
                pdlg.contents.lpVtbl.contents.SetDefaultExtension(pdlg, "")

            try:
                pdlg.contents.lpVtbl.contents.Show(pdlg, hwnd)
            except OSError:
                return None
            with _get_pIShellItemArray() as psia:
                pdlg.contents.lpVtbl.contents.GetResults(pdlg, byref(psia))
                count = c_uint32()
                psia.contents.lpVtbl.contents.GetCount(psia, byref(count))
                ret = []
                for i in range(count.value):
                    with _get_pIShellItem() as pitem:
                        psia.contents.lpVtbl.contents.GetItemAt(psia, i, byref(pitem))
                        p_path = c_wchar_p()
                        pitem.contents.lpVtbl.contents.GetDisplayName(
                            pitem, SIGDN_FILESYSPATH, byref(p_path)
                        )
                        ret.append(p_path.value)
                        oledll.ole32.CoTaskMemFree(p_path)
    return ret


def asksaveasfilename(
    hwnd: c_void_p = None,
    initialdir: str = None,
    initialfile: str = None,
    filetypes: List[Tuple[str, str]] = None,
) -> str:
    """Ask for a file to save as, and return the path.

    Args:
        hwnd: A handle to the window that owns the dialog box.
        initialdir: The directory that the dialog starts in.
        initialfile: The file name used to initialize the File Name edit control.
        filetypes: The file types that the dialog can save.

    Returns:
        The file path or None if cancel button selected.
    """
    with _COM():
        with _get_pIFileSaveDialog() as pdlg:
            if initialdir:
                with _get_pIShellItem() as pitem:
                    oledll.shell32.SHCreateItemFromParsingName(
                        initialdir, None, IID_IShellItem, byref(pitem)
                    )
                    pdlg.contents.lpVtbl.contents.SetFolder(pdlg, pitem)

            if initialfile:
                pdlg.contents.lpVtbl.contents.SetFileName(pdlg, initialfile)

            if filetypes:
                rgSpec = (COMDLG_FILTERSPEC * len(filetypes))()
                for i, ft in enumerate(filetypes):
                    rgSpec[i] = COMDLG_FILTERSPEC.from_tuple(ft)
                pdlg.contents.lpVtbl.contents.SetFileTypes(pdlg, len(filetypes), rgSpec)
                pdlg.contents.lpVtbl.contents.SetDefaultExtension(pdlg, "")

            try:
                pdlg.contents.lpVtbl.contents.Show(pdlg, hwnd)
            except OSError:
                return None
            with _get_pIShellItem() as pItem:
                pfilepath = c_wchar_p()
                pdlg.contents.lpVtbl.contents.GetResult(pdlg, byref(pItem))
                pItem.contents.lpVtbl.contents.GetDisplayName(
                    pItem, SIGDN_FILESYSPATH, byref(pfilepath)
                )
                ret = pfilepath.value
                oledll.ole32.CoTaskMemFree(pfilepath)
    return ret

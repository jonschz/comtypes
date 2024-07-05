from ctypes import Array, c_ubyte, c_ulong, HRESULT, POINTER, pointer
from typing import Tuple, TYPE_CHECKING

from comtypes import COMMETHOD, GUID, IUnknown


class ISequentialStream(IUnknown):
    """Defines methods for the stream objects in sequence."""

    _iid_ = GUID("{0C733A30-2A1C-11CE-ADE5-00AA0044773D}")
    _idlflags_ = []

    _methods_ = [
        # Note that these functions are called `Read` and `Write` in Microsoft's documentation,
        # see https://learn.microsoft.com/en-us/windows/win32/api/objidl/nn-objidl-isequentialstream.
        # However, the comtypes code generation detects these as `RemoteRead` and `RemoteWrite`
        # for very subtle reasons, see e.g. https://stackoverflow.com/q/19820999/. We will not
        # rename these in this manual import for the sake of consistency.
        COMMETHOD(
            [],
            HRESULT,
            "RemoteRead",
            # This call only works if `pv` is pre-allocated with `cb` bytes,
            # which cannot be done by the high level function generated by metaclasses.
            # Therefore, we override the high level function to implement this behaviour
            # and then delegate the call the raw COM method.
            (["out"], POINTER(c_ubyte), "pv"),
            (["in"], c_ulong, "cb"),
            (["out"], POINTER(c_ulong), "pcbRead"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "RemoteWrite",
            (["in"], POINTER(c_ubyte), "pv"),
            (["in"], c_ulong, "cb"),
            (["out"], POINTER(c_ulong), "pcbWritten"),
        ),
    ]

    def RemoteRead(self, cb: int) -> Tuple["Array[c_ubyte]", int]:
        """Reads a specified number of bytes from the stream object into memory
        starting at the current seek pointer.
        """
        # Behaves as if `pv` is pre-allocated with `cb` bytes by the high level func.
        pv = (c_ubyte * cb)()
        pcb_read = pointer(c_ulong(0))
        self.__com_RemoteRead(pv, c_ulong(cb), pcb_read)  # type: ignore
        # return both `out` parameters
        return pv, pcb_read.contents.value

    if TYPE_CHECKING:

        def RemoteWrite(self, pv: "Array[c_ubyte]", cb: int) -> int:
            """Writes a specified number of bytes into the stream object starting at
            the current seek pointer.
            """
            ...


# fmt: off
__known_symbols__ = [
    'ISequentialStream',
]
# fmt: on

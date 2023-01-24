import unittest as ut

from ctypes import POINTER, byref, c_ulong, c_bool, c_ubyte
import comtypes
import comtypes.client

comtypes.client.GetModule("portabledeviceapi.dll")
from comtypes.gen.PortableDeviceApiLib import IStream

class Test_IStream(ut.TestCase):
    def test_istream(self):

        # Create an IStream
        stream: IStream = POINTER(IStream)() # type: ignore
        comtypes._ole32.CreateStreamOnHGlobal(
            None,
            c_bool(True),
            byref(stream)
        )

        test_data = "Some data".encode('utf-8')
        pv = (c_ubyte * len(test_data)).from_buffer(bytearray(test_data))
        
        written = stream.RemoteWrite(pv, c_ulong(len(test_data)))
        self.assertEqual(written, len(test_data))
        # make sure the data actually gets written before trying to read back
        stream.Commit(0)

        # Move the stream back to the beginning
        STREAM_SEEK_SET = 0
        stream.RemoteSeek(0, STREAM_SEEK_SET)

        read_buffer_size = 1024

        # Option 1: Reading from IStream using the high-level method
        read_buffer, data_read = stream.RemoteRead(read_buffer_size)

        # # Option 2: Reading from IStream using the barebones COM method
        # pcb_read = pointer(c_ulong(0))
        # read_buffer = (c_ubyte * read_buffer_size)()
        # hres = stream._ISequentialStream__com_RemoteRead(read_buffer, read_buffer_size, pcb_read)
        # data_read = pcb_read.contents.value
        # assert hres == 0

        # Verification
        self.assertEqual(data_read, len(test_data))
        self.assertEqual(bytearray(read_buffer)[0:data_read], test_data)


if __name__ == "__main__":
    ut.main()

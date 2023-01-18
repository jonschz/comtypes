from __future__ import annotations # so we don't crash when the generated files don't exist

# import ctypes
from ctypes import POINTER, _Pointer, pointer, cast, c_uint32, c_uint16, c_uint, c_int8, c_ubyte, c_ulong, c_wchar_p
import unittest as ut

import comtypes
import comtypes.client

comtypes.client.GetModule("portabledeviceapi.dll")
import comtypes.gen.PortableDeviceApiLib as port_api

comtypes.client.GetModule("portabledevicetypes.dll")
import comtypes.gen.PortableDeviceTypesLib as port_types


def newGuid(*args: int) -> comtypes.GUID:
    guid = comtypes.GUID()
    guid.Data1 = c_uint32(args[0])
    guid.Data2 = c_uint16(args[1])
    guid.Data3 = c_uint16(args[2])
    for i in range(8):
        guid.Data4[i] = c_int8(args[3 + i])
    return guid


def PropertyKey(*args: int) -> _Pointer[port_api._tagpropertykey]:
    assert len(args) == 12
    assert all(isinstance(x, int) for x in args)
    propkey = port_api._tagpropertykey()
    propkey.fmtid = newGuid(*args[0:11])
    propkey.pid = c_ulong(args[11])
    return pointer(propkey)


class Test_IPortableDevice(ut.TestCase):
    # To avoid damaging or changing the environment, do not CREATE, DELETE or UPDATE!
    # Do READ only!
    def setUp(self):
        info = comtypes.client.CreateObject(
            port_types.PortableDeviceValues().IPersist_GetClassID(),
            clsctx=comtypes.CLSCTX_INPROC_SERVER,
            interface=port_types.IPortableDeviceValues,
        )
        mng = comtypes.client.CreateObject(
            port_api.PortableDeviceManager().IPersist_GetClassID(),
            clsctx=comtypes.CLSCTX_INPROC_SERVER,
            interface=port_api.IPortableDeviceManager,
        )
        p_id_cnt = pointer(c_ulong())
        mng.GetDevices(POINTER(c_wchar_p)(), p_id_cnt)
        if p_id_cnt.contents.value == 0:
            self.skipTest("There is no portable device in the environment.")
        dev_ids = (c_wchar_p * p_id_cnt.contents.value)()
        mng.GetDevices(cast(dev_ids, POINTER(c_wchar_p)), p_id_cnt)
        self.device = comtypes.client.CreateObject(
            port_api.PortableDevice().IPersist_GetClassID(),
            clsctx=comtypes.CLSCTX_INPROC_SERVER,
            interface=port_api.IPortableDevice,
        )
        self.device.Open(list(dev_ids)[0], info)

    def test_EnumObjects(self):
        WPD_OBJECT_NAME = PropertyKey(
            0xEF6B490D, 0x5CD8, 0x437A, 0xAF, 0xFC, 0xDA, 0x8B, 0x60, 0xEE, 0x4A, 0x3C, 4
        )
        WPD_OBJECT_CONTENT_TYPE = PropertyKey(
            0xEF6B490D, 0x5CD8, 0x437A, 0xAF, 0xFC, 0xDA, 0x8B, 0x60, 0xEE, 0x4A, 0x3C, 7
        )
        WPD_RESOURCE_DEFAULT = PropertyKey(
            0xE81E79BE, 0x34F0, 0x41BF, 0xB5, 0x3F, 0xF1, 0xA0, 0x6A, 0xE8, 0x78, 0x42, 0
        )
        folderType = newGuid(
            0x27E2E392, 0xA111, 0x48E0, 0xAB, 0x0C, 0xE1, 0x77, 0x05, 0xA0, 0x5F, 0x85
        )
        functionalType = newGuid(
            0x99ED0160, 0x17FF, 0x4C44, 0x9D, 0x98, 0x1D, 0x7A, 0x6F, 0x94, 0x19, 0x21
        )
        
        content = self.device.Content()
        properties = content.Properties()
        propertiesToRead = comtypes.client.CreateObject(
            port_types.PortableDeviceKeyCollection,
            clsctx=comtypes.CLSCTX_INPROC_SERVER,
            interface=port_api.IPortableDeviceKeyCollection,
        )
        propertiesToRead.Add(WPD_OBJECT_NAME)
        propertiesToRead.Add(WPD_OBJECT_CONTENT_TYPE)
        def downloadFirstFile(objectID):
            # print(contentID)
            values = properties.GetValues(objectID, propertiesToRead)
            contenttype = values.GetGuidValue(WPD_OBJECT_CONTENT_TYPE)
            is_folder = contenttype in [folderType, functionalType]
            print(f"{values.GetStringValue(WPD_OBJECT_NAME)} ({'folder' if is_folder else 'file'})")
            if is_folder:
                # traverse into the children
                enumobj = content.EnumObjects(c_ulong(0), objectID, None)
                for x in enumobj:
                    if downloadFirstFile(x):
                        return True
                return False
            else:
                resources = content.Transfer()
                STGM_READ = c_uint(0)
                optimalTransferSizeBytes = pointer(c_ulong(0))
                pFileStream = POINTER(port_api.IStream)()
                optimalTransferSizeBytes, pFileStream = resources.GetStream(
                    objectID,
                    WPD_RESOURCE_DEFAULT,
                    STGM_READ,
                    optimalTransferSizeBytes,
                )
                blockSize = optimalTransferSizeBytes.contents.value
                fileStream = pFileStream.value
                while True:
                    #
                    # # This crashes without an exception - likely a segfault somewhere deeper.
                    # # You won't see "Read data: ..." or "Download complete"
                    #
                    # buf, data_read = fileStream.RemoteRead(c_ulong(blockSize))
                    #
                    # This works:
                    # (it is cleaner to pull the first two steps outside of the loop)
                    buf = (c_ubyte * blockSize)()
                    pdata_read = pointer(c_ulong(0))
                    fileStream._ISequentialStream__com_RemoteRead(buf, c_ulong(blockSize), pdata_read)
                    data_read = pdata_read.contents.value
                    
                    print(f"Read data: {data_read}")
                    if data_read == 0:
                        break
                    # optional: save the file locally
                    # with open('firstfile.png', 'wb') as outputStream:
                    #     outputStream.write(bytearray(buf)[0:data_read])
                print("Download complete")
                # stop after reading the first file. If the wrong code doesn't crash, set this to False
                return True
            
        downloadFirstFile("DEVICE")


if __name__ == "__main__":
    ut.main()

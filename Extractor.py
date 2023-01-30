import olefile
import os
import zlib

DUMP_PATH = "./Dump/"
def doDecompress(bin_data):
	zobj = zlib.decompressobj(-zlib.MAX_WBITS)
	bin_data = zobj.decompress(bin_data)

	return bin_data

def FindBinData(hwp_struct, Type):
    ret_array = list()
    for i in range(len(hwp_struct)):
        if hwp_struct[i][0] == "BinData":
            name = hwp_struct[i][1].upper()
            fn = name.find(Type)
            if fn != -1:
                tmp = hwp_struct[i][0]
                tmp += '/'
                tmp += hwp_struct[i][1] # PostScript Name
                ret_array.append(tmp)
    return ret_array

def ShowTree(arr):
    print("■ BinData")
    for name in arr:
        print("   └ %s" % name)
                                
def Init(hwpfile):
    BIN_ARRAY = []
    hwp_struct = hwpfile.listdir()
    
    PS_ARRAY = FindBinData(hwp_struct, ".PS")
    EPS_ARRAY = FindBinData(hwp_struct, ".EPS")
    OLE_ARRAY = FindBinData(hwp_struct, ".OLE")
    BIN_ARRAY.extend(PS_ARRAY)
    BIN_ARRAY.extend(EPS_ARRAY)
    BIN_ARRAY.extend(OLE_ARRAY)
    
    if BIN_ARRAY == []:
        print("Not Found!")
        return 1
    
    ShowTree(BIN_ARRAY)
    if not os.path.isdir(DUMP_PATH):
        os.mkdir(DUMP_PATH)
    for i in range(len(BIN_ARRAY)):
        bin_stream = hwpfile.openstream(BIN_ARRAY[i])
        bin_data = bin_stream.read()
        bin_data = doDecompress(bin_data)
        
        outfile_name = BIN_ARRAY[i].split('/')
        outfile = open(DUMP_PATH + outfile_name[1], 'wb')
        outfile.write(bin_data)
        outfile.close()
    return 0


filepath = input("Path : ")
hwpfile = olefile.OleFileIO(filepath)
Init(hwpfile)

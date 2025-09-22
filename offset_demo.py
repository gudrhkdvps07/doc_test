import olefile
import struct

doc_path = "./이름 테스트.doc"
ole = olefile.OleFileIO(doc_path)

with ole.openstream("WordDocument") as stream:
    data = stream.read()

print("WordDocument 스트림 크기:", len(data))

#FIB에서 fcClx, lcbClx읽기
fcClx = struct.unpack_from("<I", data, 0x01A2)[0]
lcbClx = struct.unpack_from("<I", data, 0x01A6)[0]

print(f"fcClx = {hex(fcClx)} ({fcClx})")
print(f"lcbClx = {hex(lcbClx)} ({lcbClx})")

#Clx 블록 추출
clx = data[fcClx:fcClx + lcbClx]
print("CLx 크기:", len(clx), "bytes")
print("Clx 시작 바이트:", clx[:16])

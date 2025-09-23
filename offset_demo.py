import olefile, struct

doc_path = "./테스트.doc"
with olefile.OleFileIO(doc_path) as ole:
    # WordDocument 스트림 읽기
    word_data = ole.openstream("WordDocument").read()

    # FIB에서 fcClx, lcbClx 읽기
    fcClx = struct.unpack_from("<I", word_data, 0x01A2)[0]
    lcbClx = struct.unpack_from("<I", word_data, 0x01A6)[0]

    # fWhichTblStm 플래그 확인
    fib_base_flags = struct.unpack_from("<H", word_data, 0x000A)[0]
    fWhichTblStm = (fib_base_flags & 0x0200) != 0
    tbl_stream = "1Table" if fWhichTblStm else "0Table"

    # Table 스트림 읽기
    table_data = ole.openstream(tbl_stream).read()

print("WordDocument 스트림 크기:", len(word_data))
print(f"fcClx = {hex(fcClx)} ({fcClx})")
print(f"lcbClx = {hex(lcbClx)} ({lcbClx})")
print(f"이 문서는 {'1Table' if fWhichTblStm else '0Table'} 스트림입니다.")
print("Table 스트림 크기:", len(table_data))

#Clx 블록 추출
if lcbClx == 0:
    raise ValueError("CLX 길이가 0입니다 (텍스트 조각 정보 없음)")
if fcClx + lcbClx > len(table_data):
    raise ValueError("CLX 범위가 테이블 스트림을 벗어납니다")
clx = table_data[fcClx:fcClx + lcbClx]
print("CLx 크기:", len(clx), "bytes")
print("Clx 시작 바이트:", clx[:16])

#CLx 안에서 PlcPcd 추출
def extract_plcpcd(clx: bytes) -> bytes:
    i = 0
    while i < len(clx):
        tag = clx[i]
        i += 1
        if tag == 0x01:  # Prc
            if i + 2 > len(clx):
                raise ValueError("잘못된 Clx: Prc 헤더가 짧음")
            cb = struct.unpack_from("<H", clx, i)[0]
            i += 2 + cb
        elif tag == 0x02:  # Pcdt
            if i + 4 > len(clx):
                raise ValueError("잘못된 Clx: Pcdt 길이 누락")
            lcb = struct.unpack_from("<I", clx, i)[0]
            i += 4
            if i + lcb > len(clx):
                raise ValueError("잘못된 Clx: PlcPcd 범위 초과")
            return clx[i:i+lcb]  # 정상 반환
        else:
            raise ValueError(f"알 수 없는 CLX 태그: 0x{tag:02X}")

    raise ValueError("Clx 안에서 Pcdt(0x02)를 찾지 못했음")

plcpcd = extract_plcpcd(clx)
print("PlcPcd 크기: ",len(plcpcd))
print(plcpcd.hex())

def parse_plcpcd(plcpcd: bytes):
    size = len(plcpcd)
    if (size - 4) % 12 != 0:
        raise ValueError("PlcPcd 길이가 예상 형식(4*(n+1)+8*n)에 맞지 않습니다")
    n = (size - 4) // 12  # size = 4*(n+1) + 8*n (n은 조각 개수)
    
    # aCp 배열 읽기
    acp = [struct.unpack_from("<I", plcpcd, 4*i)[0] for i in range(n+1)]
    
    #PCD 배열 시작 위치
    pcd_off = 4 * (n+1)
    
    pieces = []
    for k in range(n):
        pcd_bytes = plcpcd[pcd_off + 8*k : pcd_off + 8*(k+1)] #PCD = 8byte
        flags = struct.unpack_from("<h", pcd_bytes, 0)[0] #앞 2바이트는 flag

        #이후 4바이트 = fc
        fc_raw = struct.unpack_from("<i", pcd_bytes, 2)[0]

        # fcRaw 해석
        fc = fc_raw & 0x3FFFFFFF  # 하위 30비트만
        fCompressed = (fc_raw & 0x40000000) != 0
        print(f"fc_raw=0x{fc_raw:08X}, fc={fc}, fCompressed={fCompressed}")
        
        prm = struct.unpack_from("<h", pcd_bytes, 0)[0]

        cp_start = acp[k]
        cp_end   = acp[k+1]
        char_count = cp_end - cp_start
        byte_count = char_count if fCompressed else char_count * 2

        pieces.append({
            "piece_index": k,
            "cp_start": cp_start,
            "cp_end": cp_end,
            "char_count": char_count,
            "flags": flags,
            "fc": fc,
            "fCompressed": fCompressed,
            "byte_count": byte_count,
            "prm": prm
        })
    
    return pieces

pieces = parse_plcpcd(plcpcd)
print("조각 개수:", len(pieces))
for p in pieces[:5]:
    print(p)

'''
def decode_piece(chunk: bytes, fCompressed: bool) -> str:
    if fCompressed:
        return chunk.decode("cp1252", errors="replace") #1 byte
    else:
        return chunk.decode("utf-16le", errors="replace")  #2 byte

#텍스트 추출
def extract_full_text(word_data: bytes, pieces):
    texts = []
    for p in pieces:
        #WordDocument에서 해당 조각의 바이트 범위 잘라오기
        chunk = word_data[p["fc"]: p["fc"] + p["byte_count"]]
        text = decode_piece(chunk, p["fCompressed"])
        texts.append(text)
    return "".join(texts)

full_text = extract_full_text(word_data, pieces)
print("추출된 텍스트: ", full_text)

replacement_text = full_text.replace("함근희", "***")
print("치환된 텍스트:", replacement_text)
'''
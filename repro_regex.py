
import re

linea = "      30/12/25  Remuneración de Saldo          0206580294                     118.932,21"
linea = linea.strip()

print(f"Linea: '{linea}'")

pattern_monto = re.compile(r"((?:\d{1,3}(?:\.\d{3})*)?,\d{2}-?)")
matches = pattern_monto.findall(linea)

print(f"Matches found: {len(matches)}")
for m in matches:
    print(f"Match: '{m}'")

if len(matches) == 1:
    monto_str_raw = matches[0]
    print(f"Single match extraction: {monto_str_raw}")
    
    # Check simple regex for saldo inicial too
    header_line = "             Saldo del período anterior                                                      96.582.602,40"
    if "Saldo del per" in header_line and "anterior" in header_line:
        print("Header detection: SUCCESS")
        match_saldo = re.search(r"([\d\.]+,\d{2}[\-]?)", header_line.strip())
        if match_saldo:
            print(f"Saldo inicial extracted: '{match_saldo.group(1)}'")
    else:
        print("Header detection: FAILED")

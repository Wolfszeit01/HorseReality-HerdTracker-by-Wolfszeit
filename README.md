Fix 1; 30.05.2026

There was an error involving missing formulas. You can easily fix this by copying the formulas form first line using the fill handle (the small square at the bottom right corner of the active cell) or recopying the sheet.

Additionally, there was another minor error when retrieving the breeds. This can also be resolved—either by recopying the entire Tracker, or by replacing the specific files
export.gs
Outsidestudsbackend.gs
with the ones attached here.

To do this, simply navigate to "Extensions" > "Apps Script" and replace the respective files.

Fix 2; 31.05.2026

Replace: 
Outsidestudsbackend.gs
outsideStudsModal.html
FoalCalulator.gs

To do this, simply navigate to "Extensions" > "Apps Script" and replace the respective files.

 for Outside Stud

Colum AI: =COUNTIF(G3:T3,"VG")
Colum AJ: =COUNTIF(G3:T3,"G+")
Colum AK: =COUNTIF(G3:T3,"G") 
Colum AL: =COUNTIF(G3:T3,"G-")
Colum AM:  =COUNTIF(G3:T3,"A")
Colum AN:  =COUNTIF(G3:T3,"BA")
Colum AO:  =COUNTIF(G3:T3,"P")

Column BP: =INDEX($BA$1:$BO$1,MATCH(MAX(BC3,BE3,BG3,BI3,BK3,BM3,BO3),BB3:BO3,0))

for Her Tracker

Colum L: 
=IFNA(
  LET(
    vater, VLOOKUP(C2, Pedigree!$B$1:G, 5, FALSE),
    linie, VLOOKUP(C2, Pedigree!$B$1:G, 3, FALSE),
    IF(AND(vater="", linie=""), "Unknown",
      IF(vater="", "Unknown", vater) & " (" & IF(linie="", "Unknown", linie) & ")"
    )
  ),
  "Unknown"
)

Colum M: 
=IFNA(
  LET(
    mutter, VLOOKUP(C2, Pedigree!$B$1:G, 6, FALSE),
    linie, VLOOKUP(C2, Pedigree!$B$1:G, 4, FALSE),
    IF(AND(mutter="", linie=""), "Unknown",
      IF(mutter="", "Unknown", mutter) & " (" & IF(linie="", "Unknown", linie) & ")"
    )
  ),
  "Unknown"
)

Colum S: 
=LET(breed, $E2, id, $C2, 
  sheet, IFS(ISNUMBER(SEARCH("Icelandic", breed)), "ICE_Horse Stats", ISNUMBER(SEARCH("Kathiawari", breed)), "KATH_Horse Stats", TRUE, "Horse Stats"),
  offset, IFS(ISNUMBER(SEARCH("Icelandic", breed)), 4, ISNUMBER(SEARCH("Kathiawari", breed)), 2, TRUE, 0),
  VLOOKUP(id, INDIRECT(sheet & "!B:BX"), 44 + offset, FALSE)
)

Colum T: =IFERROR(
  TRIM(
    LET(
      breed, $E2,
      horseID, $C2,
      targetSheet, IFS(
        ISNUMBER(SEARCH("Icelandic", breed)), "ICE_Horse Stats",
        ISNUMBER(SEARCH("Kathiawari", breed)), "KATH_Horse Stats",
        TRUE, "Horse Stats"
      ),
      offset, IFS(
        ISNUMBER(SEARCH("Icelandic", breed)), 4,
        ISNUMBER(SEARCH("Kathiawari", breed)), 2,
        TRUE, 0
      ),
      rangeID, INDIRECT(targetSheet & "!$B:$B"),
      rowIdx, MATCH(horseID, rangeID, 0),
      valVG,    INDEX(INDIRECT(targetSheet & "!1:" & rowIdx), rowIdx, 28 + offset),
      valGplus, INDEX(INDIRECT(targetSheet & "!1:" & rowIdx), rowIdx, 29 + offset),
      valG,     INDEX(INDIRECT(targetSheet & "!1:" & rowIdx), rowIdx, 30 + offset),
      valGminus,INDEX(INDIRECT(targetSheet & "!1:" & rowIdx), rowIdx, 31 + offset),
      valA,     INDEX(INDIRECT(targetSheet & "!1:" & rowIdx), rowIdx, 32 + offset),
      valBA,    INDEX(INDIRECT(targetSheet & "!1:" & rowIdx), rowIdx, 33 + offset),
      IF(valVG<>"", valVG & " VG ", "") &
      IF(valGplus<>"", valGplus & " G+ ", "") &
      IF(valG, valG & " G ", "") &
      IF(valGminus<>"", valGminus & " G- ", "") &
      IF(valA<>"", valA & " A ", "") &
      IF(valBA<>"", valBA & " BA ", "")
    )
  ), 
  "not found"
)


Colum W:
=LET(breed, $E2, id, $C2,
  sheet, IFS(ISNUMBER(SEARCH("Icelandic", breed)), "ICE_Horse Stats", ISNUMBER(SEARCH("Kathiawari", breed)), "KATH_Horse Stats", TRUE, "Horse Stats"),
  off, IFS(ISNUMBER(SEARCH("Icelandic", breed)), 4, ISNUMBER(SEARCH("Kathiawari", breed)), 1, TRUE, 0),
  VLOOKUP(id, INDIRECT(sheet & "!B:BW"), 64 + off, 0) & " | " & 
  VLOOKUP(id, INDIRECT(sheet & "!B:BW"), 65 + off, 0) & " | " & 
  VLOOKUP(id, INDIRECT(sheet & "!B:BW"), 66 + off, 0) & " | " & 
  VLOOKUP(id, INDIRECT(sheet & "!B:BW"), 67 + off, 0)
)

Colum X: 
=LET(
  breed, $E2, 
  id, $C2, 
  targetSheet, IFS(
    ISNUMBER(SEARCH("Icelandic", breed)), "ICE_Horse Stats", 
    ISNUMBER(SEARCH("Kathiawari", breed)), "KATH_Horse Stats", 
    TRUE, "Horse Stats"
  ),
  offset, IFS(
    ISNUMBER(SEARCH("Icelandic", breed)), 4, 
    ISNUMBER(SEARCH("Kathiawari", breed)), 2, 
    TRUE, 0
  ),
  VLOOKUP(id, INDIRECT(targetSheet & "!B:BW"), 59 + offset, FALSE)
)

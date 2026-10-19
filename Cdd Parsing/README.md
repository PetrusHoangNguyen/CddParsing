# Cdd Parsing

`cdd_parser.py` reads a CDD file exported from Vector CANdela Studio and writes its diagnostic data (services, DIDs, DTCs and RIDs) to either an XML file or an Excel workbook. It needs Python 3.9 or newer. XML output uses only the standard library; Excel output also needs `openpyxl` (`pip install -r requirements.txt`).

```
# XML (default)
python cdd_parser.py ../RADAR_G7_CHERY03_APP_and_BOOT.cdd -o RADAR_G7_CHERY03_APP_and_BOOT.xml

# Excel
python cdd_parser.py ../RADAR_G7_CHERY03_APP_and_BOOT.cdd --format excel -o RADAR_G7_CHERY03_APP_and_BOOT.xlsx
```

| Option | Meaning |
|--------|---------|
| `-f`, `--format` | `xml` (default) or `excel` |
| `-o`, `--output` | output file; without it the file is written next to the CDD as `<name>.xml` or `<name>.xlsx` |

## Excel layout

| Sheet   | One row per | Main columns |
|---------|-------------|--------------|
| Service | service of a diagnostic instance, per variant | SID, request bytes, addressing, sessions, security |
| DID     | data element of a DID | DID, variants, length, element position and bit length, data type details, read/write access |
| DTC     | DTC record of a fault memory | DTC, SAE code (e.g. `U0418-81`), description, all DTC record fields |
| RID     | routine of `$31 RoutineControl` | RID, sub-functions with request bytes, sessions, security, parameters |

Both formats carry the same data: the Excel sheets are built from the XML tree. The last columns of each sheet (`CDD Ref`, `DID Ref`, `Element Ref`) link each row to its element in the CDD; `excel_to_cdd.py` needs them, so leave them as they are.

## Editing the Excel file and writing a new CDD

`excel_to_cdd.py` takes the original CDD and an Excel file exported from it, and writes a new CDD with your edits:

```
python cdd_parser.py ../RADAR_G7_CHERY03_APP_and_BOOT.cdd --format excel -o export.xlsx
# edit export.xlsx
python excel_to_cdd.py ../RADAR_G7_CHERY03_APP_and_BOOT.cdd export.xlsx -o RADAR_G7_CHERY03_edited.cdd --report report.txt
```

The original CDD is the template. Only the values you changed are rewritten in place; everything else, including ids, references, quotes and line endings, stays byte for byte the same. Running it on an unedited export gives a file identical to the original.

| Sheet | Editable columns |
|-------|------------------|
| Service | Sessions, Security, Addressing |
| DID | DID, Name, Qualifier, Read/Write Sessions and Security, Data Element, Element Qualifier, Conversion, Unit, Value Table (texts only) |
| DTC | DTC or DTC Code, Description, and every DTC record field the DTC already has |
| RID | RID, Name, Qualifier, Description, Sessions, Security |

All other columns are computed from the CDD and are read-only. Adding or deleting rows is not supported yet.

How some edits are applied:

- **Sessions / Security** take state names separated by commas, `All`, or `-` for none. A routine's sessions and security apply to all its sub-functions.
- **DID number** also updates the identifier of the read, write and IO control services that use the DID.
- **Conversion, Unit and Value Table** belong to the data type, which can be shared with other DIDs. Other rows that change with it are listed as warnings.
- Services shared between variants (for example Tester Present) change in every variant that uses them.

The script prints a report. Errors name the sheet, row and column, for example an unknown session name, a value that is not hex, a duplicate DID or DTC number, an edit to a read-only column, or a new or deleted row. **If there is any error, no CDD is written.** After writing, the script reads the new CDD back and checks that every edit is in it. Warnings list other rows that changed as a side effect, and `CHANGED` lines list each applied edit.

CANdela Studio itself is not run by the script: open the new CDD in CANdela to confirm the import.

## XML layout

```xml
<CddExport source="..." ecu="...">
  <Services count="...">
    <Variant name="APP">
      <Service sid="22" request="22 F190" diagClass="Identification" instance="VIN Data Identifier"
               name="Read" addressing="Physical" sessions="..." security="..." />
  <DIDs count="...">
    <DID id="E681" name="..." variants="APP" length="2">
      <Read sid="22" sessions="..." security="..." />
      <DataElement name="Vehicle Speed" startByte="0" startBit="0" bitLength="16" dataType="Linear: Vehicle Speed"
                   codedType="Unsigned 16 bit" byteOrder="Big endian (Motorola)" physicalType="Float64 64 bit"
                   displayFormat="Float" unit="km/h">
        <Conversion raw="0..8190" factor="0.0625" offset="0" />
        <Value raw="0">Off</Value>              <!-- text tables -->
        <Invalid raw="128..255">invalidSignal</Invalid>
  <DTCs count="...">
    <FaultMemory variant="APP" name="Primary Fault Memory">
      <DTC id="570016" code="C1700-16" description="...">
        <Field name="DS_DTCPriority">4</Field>
  <RIDs count="...">
    <Variant name="Boot">
      <RID id="DD04" name="Version switchover" sessions="..." security="...">
        <Request subFunction="Start">31 01 DD04</Request>
        <Parameters group="Control Options (Start)"><Parameter name="..." /></Parameters>
```

Empty attributes are left out.

## Notes

- **request** shows the constant leading bytes of the request: SID, sub-function and identifier (for example `31 01 0531`). Variable data that follows is not shown.
- **sessions / security** come from the `mayBeExec` attribute, which lists the states of all state groups in document order. `All` means the CDD sets no restriction. `-` means the CDD lists no state of that group for the service.
- DTCs of the shared record pool, which belong to no variant, are listed under `variant="Record pool (no variant)"`.
- **DataElement** `startByte` / `startBit` count from the first data byte after the DID.
- **Conversion** lists the factor, divisor and offset as stored in the CDD. The physical value is assumed to be `raw x factor / divisor + offset`; this formula is inferred, not stated in the file.
- **byteOrder** maps CANdela `bo="21"` to big endian (Motorola) and `bo="12"` to little endian (Intel).

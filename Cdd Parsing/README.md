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

Both formats carry the same data: the Excel sheets are built from the XML tree.

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

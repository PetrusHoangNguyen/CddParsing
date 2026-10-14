# Cdd Parsing

`cdd_parser.py` reads a CDD file exported from Vector CANdela Studio and writes its diagnostic data to one easy-to-read XML file. It needs only Python 3.9 or newer, with no extra packages.

```
python cdd_parser.py ../RADAR_G7_CHERY03_APP_and_BOOT.cdd -o RADAR_G7_CHERY03_APP_and_BOOT.xml
```

Without `-o`, the XML is written next to the CDD file with the `.xml` extension.

## Output layout

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

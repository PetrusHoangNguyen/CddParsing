# Cdd Parsing

`cdd_parser.py` reads a CDD file exported from Vector CANdela Studio and writes its diagnostic data to an Excel workbook with four sheets:

| Sheet   | One row per | Main columns |
|---------|-------------|--------------|
| Service | service of a diagnostic instance, per variant | SID, request bytes, addressing, sessions, security |
| DID     | data identifier in the `DIDS` list | DID, variants, length, data elements, read/write access |
| DTC     | DTC record of a fault memory | DTC, SAE code (e.g. `U0418-81`), description, all DTC record fields |
| RID     | routine of `$31 RoutineControl` | RID, sub-functions with request bytes, sessions, security, parameters |

## Usage

```
pip install -r requirements.txt
python cdd_parser.py ../RADAR_G7_CHERY03_APP_and_BOOT.cdd -o RADAR_G7_CHERY03_APP_and_BOOT.xlsx
```

Without `-o`, the workbook is written next to the CDD file with the `.xlsx` extension.

## Notes

- **Request** shows the constant leading bytes of the request: SID, sub-function and identifier (for example `31 01 0531`). Variable data that follows is not shown.
- **Sessions / Security** come from the `mayBeExec` attribute, which lists the states of all state groups in document order. `All` means the CDD sets no restriction. `-` means the CDD lists no state of that group for the service.
- DTCs of the shared record pool, which belong to no variant, are listed with the variant `Record pool (no variant)`.

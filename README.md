# mdbToXML

A fast, streaming CLI tool that converts Microsoft Access `.mdb` / `.accdb` databases to XML and XSD files — compatible with the format produced by Access's native `ExportXML`.

> **Windows-only** — requires the Microsoft ACE OLEDB 12.0 provider.

## Features

- 🚀 **Streaming architecture** — handles databases over 1 GB without loading everything into memory
- 📐 **Access-native XML/XSD format** — generates `od:`, `xsi:` namespaces identical to Access `ExportXML`
- ✂️ **File splitting** — split large exports into multiple files by row count (e.g. every 10,000 rows)
- 🛡️ **Corruption-resilient** — gracefully skips corrupted OLE Object fields and damaged rows instead of crashing
- 📦 **Zero dependencies** — ships as a single self-contained `.exe`, no .NET runtime required on the target machine
- 🔄 **xml2js compatible** — always writes full end tags (`<tag></tag>`) for seamless JavaScript/Node.js parsing
- ⚡ **Parallel table export** — exports multiple tables concurrently when not using `--single-file` mode
- 🔍 **Rich schema extraction** — reads primary keys, indexes, foreign keys, and DAO field-level properties

## Quick Start

### Download

Grab the latest `mdbToXML.exe` from the [Releases](../../releases) page. No installation needed.

### Basic Usage

```bash
# Export all tables to XML + XSD (output alongside the .mdb file)
mdbToXML export "C:\Path\To\Database.mdb"

# Export to a specific output directory with verbose progress
mdbToXML export "C:\Path\To\Database.mdb" -o ./output -v

# Inspect database structure without exporting
mdbToXML info "C:\Path\To\Database.mdb"
```

### Split Large Exports

When dealing with large tables, use `--split-size` to break the output into manageable chunks.  
If total rows are fewer than the split size, **no sequence suffix is added** — you get a single clean file.

```bash
# Split every 10,000 rows → Table_001.xml, Table_002.xml, ...
mdbToXML export "C:\Path\To\Database.mdb" -o ./output --split-size 10000

# Only splits if needed; 8,000 rows with --split-size 10000 → Table.xml (no suffix)
```

### Export Specific Tables

```bash
mdbToXML export "C:\Path\To\Database.mdb" -t "Customers,Orders" -o ./output -v
```

### Merge All Tables into One File

```bash
mdbToXML export "C:\Path\To\Database.mdb" --single-file -o ./output
```

## Commands

| Command  | Description                                      |
| -------- | ------------------------------------------------ |
| `info`   | Analyze and display database structure            |
| `export` | Export tables to XML and/or XSD files             |

## Export Options

| Option               | Description                                            |
| -------------------- | ------------------------------------------------------ |
| `-o, --output <dir>` | Output directory (default: same directory as the .mdb)  |
| `-t, --tables <t1,t2>` | Export specific tables only (comma-separated)        |
| `--single-file`      | Merge all tables into a single XML/XSD file             |
| `--split-size <num>` | Split output into new files every `<num>` rows          |
| `--xsd-only`         | Generate XSD schema only, skip data export              |
| `--xml-only`         | Generate XML data only, skip XSD generation             |
| `--simple-xsd`       | Generate simplified XSD without DAO metadata            |
| `--batch-size <num>` | Rows to buffer before flushing to disk (default: 10000) |
| `-v, --verbose`      | Show real-time progress with row counts and speed       |

## Output Format

The generated XML follows the Microsoft Access `ExportXML` format:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<dataroot xmlns:od="urn:schemas-microsoft-com:officedata"
          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
          xsi:noNamespaceSchemaLocation="TableName.xsd"
          generated="2026-07-06T15:00:00">
  <TableName>
    <Column1>value</Column1>
    <Column2></Column2>
  </TableName>
</dataroot>
```

The companion XSD includes Access-specific annotations (`od:jetType`, `od:sqlSType`, `od:nonNullable`, etc.) for full schema fidelity.

## Prerequisites

- **Windows x64** (Windows 10 / 11 / Server 2016+)
- **Microsoft Access Database Engine 2010 Redistributable** (or later)  
  Download: [AccessDatabaseEngine_X64.exe](https://www.microsoft.com/en-us/download/details.aspx?id=54920)

> **Note:** If you already have Microsoft Office (32-bit or 64-bit) installed, the ACE provider is typically already available. The tool requires the **x64** version of the provider.

## Building from Source

```bash
# Clone the repository
git clone https://github.com/your-username/mdbToXML.git
cd mdbToXML

# Build (debug)
dotnet build

# Run directly
dotnet run -- export "C:\Path\To\Database.mdb" -o ./output -v

# Publish as a self-contained single executable
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

The compiled executable will be at:  
`bin/Release/net8.0-windows/win-x64/publish/mdbToXML.exe`

## Project Structure

```
mdbToXML/
├── Core/
│   ├── AccessDatabase.cs        # OleDb connection & schema queries
│   ├── DataExporter.cs          # Streaming XML writer with split support
│   ├── FieldPropertyReader.cs   # DAO COM interop for rich metadata
│   ├── SchemaExtractor.cs       # Table/column/index schema extraction
│   └── XsdGenerator.cs          # Access-compatible XSD generation
├── Models/
│   ├── ColumnInfo.cs            # Column metadata record
│   ├── FieldPropertyInfo.cs     # DAO field property record
│   ├── IndexInfo.cs             # Index metadata record
│   ├── RelationshipInfo.cs      # Foreign key relationship record
│   ├── TablePropertyInfo.cs     # Table-level property record
│   └── TableSchema.cs           # Aggregate table schema record
├── Utils/
│   ├── ProgressReporter.cs      # Console progress & summary display
│   ├── TypeMapper.cs            # OleDb → XSD type mapping & formatting
│   └── XmlSanitizer.cs          # XML-safe element names & values
├── exampleFile/                 # Sample XML/XSD output for reference
├── Program.cs                   # CLI entry point & argument parser
├── mdbToXML.csproj              # Project configuration
└── NuGet.Config                 # Package source configuration
```

## Handling Corrupted Databases

Real-world Access databases often contain corrupted OLE Object fields or damaged rows. This tool handles them gracefully:

- **Corrupted field values** — individual fields that throw read errors are exported as empty elements (`<Field></Field>`)
- **Damaged rows** — if `reader.Read()` fails mid-table, the tool stops that table's export, closes the XML with a valid `</dataroot>` tag, and continues with remaining tables
- **Progress is never lost** — all rows successfully read before a failure are preserved in the output

## License

[MIT](LICENSE)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

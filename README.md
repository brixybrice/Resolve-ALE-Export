# DaVinci Resolve – Timeline ALE & Metadata Exporter

## Description
A DaVinci Resolve **Workflow Integration Plugin** written in Python that exports the **current timeline** to:

- **ALE** with advanced post-processing
- **CSV Edit Index**
- **CSV Metadata** containing all available metadata for the clips actually used in the timeline

This tool is designed for **editorial, DIT and post-production workflows** where ALE conformity, slate normalization and metadata completeness are critical.

The plugin provides a native Resolve UI accessible from  
**Workspace → Workflow Integration**.

---

## Features

### Export formats
- **ALE**
- **CSV Edit Index**
- **CSV Metadata (timeline clips only)**
- **ALE + CSV Metadata**

### Track handling
- **Video only**  
  Forces the `Tracks` column in ALE to `V`
- **Video + Audio**

### Slate processing
- Merge slating information into a **Slate** column
- Multiple slate formats:
  - `Scene/Shot-Take`
  - `Scene/Take`
  - `Scene-Take`
- Optional additions:
  - `*` if the take is selected
  - Camera identifier
- Optional replacement of the **Name** column by the computed Slate

### Tape column control
- As set in project
- Reel
- Clipname without extension (from UNC / source path)
- Empty

### Metadata CSV export
- Exports **maximum metadata available via Resolve API**
- Limited strictly to **clips used in the current timeline**
- Video-only or Video+Audio aware

### UI
- Fixed window height
- Persistent user settings saved locally
- Native Resolve Workflow Integration UI

---

## Requirements

- **DaVinci Resolve 18 or higher**
- **Python 3.6+** bundled with Resolve
- No external Python libraries required

---

## Installation

Copy the script file into the Workflow Integration Plugins folder and restart DaVinci Resolve.

### macOS
/Library/Application Support/Blackmagic Design/DaVinci Resolve/Workflow Integration Plugins/

### Windows
%PROGRAMDATA%\Blackmagic Design\DaVinci Resolve\Support\Workflow Integration Plugins\

After restarting Resolve, the script is available in:
Workspace → Workflow Integration

---

## Usage

1. Open a project and select a timeline
2. Launch the script from **Workspace → Workflow Integration**
3. Choose:
   - Export folder
   - Export format
   - Track mode
   - Slate options
   - Tape column behavior
4. Click **Start**
5. Files are generated in the chosen directory

---

## Limitations

- Resolve does not expose all internal metadata fields via the scripting API  
  The CSV Metadata export is as exhaustive as the API allows
- UI cannot be fully locked during execution due to Resolve API limitations
- Automatic project backups during execution may interrupt exports

These are known limitations of Resolve scripting and not of this tool.

---

## License

MIT License  
Free to use, modify and redistribute.

---

## Credits

Inspired by and structurally based on the Resolve Workflow Integration approach used in  
**Resolve Stills Markers Exporter**  [oai_citation:0‡README.md](sediment://file_000000003438722f95e1791b6031010a)

Extended and adapted for editorial ALE and metadata workflows.

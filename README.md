# Orion – Batch Oriented Core Structural Calculator

**Version 3.1 · 2026 · Developer: Dr. L. Hulijeli**

Orion calculates the true orientation (dip, dip direction, strike) of planar structures measured on oriented drill core, for many measurements at once. Upload your collar, survey and structure tables (CSV or XLSX), map the columns, and Orion calculates every structure and shows the results in a compass preview and an interactive 3D viewer.

It is a batch version of the single-measurement [Oriented Core Structural Calculator](https://corestructure.fieldgeo.app/) workflow, designed for geologists working with whole drilling programmes.

## Features

- **Step 1 – Upload** Collar, Survey and Structures tables (`.csv` / `.xlsx`).
- **Step 2 – Map columns** with auto-detection and warnings for mis-mapped columns (e.g. a text column mapped to depth).
  - Collar: Hole ID, X, Y, Z, max depth
  - Survey: Hole ID, X, Y, Z, azimuth, dip, depth, hole size
  - Structures: depth (top/bottom), ori-line, type, core length/thickness, alpha, beta, gamma, vein minerals and %, vein style, comment
- **Step 3 – Results**
  - Results table with CSV / XLSX download
  - Compass & plane orientation preview
  - Interactive 3D drillhole and structure viewer:
    - transparent planes, two-tone (light = hanging wall, dark = foot wall); solid border on the hanging-wall side, dotted on the foot-wall side
    - colour by category, transparency, plane size, hole filter and depth interval
    - interactive legend (click a symbol to show/hide those structures)
    - labelled X / Y / Z axes with automatic tick spacing (toggleable)
    - Reset view, North-up plan, maximise, export to PNG / JPEG / PDF
- **Stereonet (new in 3.0):** equal-area lower-hemisphere projection of the structures currently shown in the 3D viewer, with great circles and/or poles, a dashed red mean plane (and mean pole), a settings toolbox (line thickness, pole size, what is drawn, grid, white background), a shared interactive legend and PNG / JPEG / PDF export
- **Stereonet (3.1 revision):** the maths now lives in `stereo.js` as pure, unit-tested functions (strike → dip direction, pole trend/plunge, Lambert equal-area projection, great circles, mean plane, density). The mean plane is calculated by aligning all pole vectors to one hemisphere, averaging and normalising them. New optional **pole-density contours** (light-blue → deep-blue shading with contour lines). Run the tests with `node tests/stereo.test.js`.
- **Click-to-locate:** click a plane in 3D or a line/pole on the stereonet and the matching table row is highlighted and scrolled into view
- **Results table (new in 3.0):** row check boxes with **Show in viewer** (limits both the 3D view and the stereonet to the checked structures), click-to-sort columns and Excel-style per-column filters (text, multiple values with `stringer | vn`, exclusions with `!bedding`, numbers `>10`, `<=5`, `10..20`). Table filters automatically limit the 3D viewer and stereonet; their own settings can only narrow that further
- Dark blue interface. Runs entirely locally – files never leave your computer.

## Controls (3D viewer)

| Action | Control |
|---|---|
| Rotate | Left-drag |
| Pan | Hold left + right mouse buttons and drag |
| Zoom | Mouse wheel |
| Select a plane | Click it |

## Running

**Windows executable:** download `Dist/2-Standalone_and_Auto_Install/Orion-v3.1.exe` and run it. It opens the app in an Edge or Chrome app window (one of them must be installed) and works offline.

**From source:** open `index.html` in a modern browser (all libraries are bundled in `vendor/`).

**Rebuilding the executable:** requires Python and PyInstaller; run `build_exe/build.ps1`.

## Conventions and assumptions

- Right-hand rule for strike. Alpha is the angle between the plane and the core axis; beta is measured clockwise from the bottom-of-hole line, looking down-hole.
- The hanging wall is taken as the upper side of a plane.
- Drillhole traces use the minimum-curvature method. Dip sign convention (negative = downward) is selectable.
- Gamma and hole size are read but not yet used in calculations.

## Installing the signed MSIX (Windows 10/11)

**Quick install:** download the contents of `Dist/2-Standalone_and_Auto_Install` (`Install-Orion.bat`, `Install-Orion.ps1`, `Orion-Certificate.cer`, `Orion-v3.1.msix`) into one folder and double-click **`Install-Orion.bat`**. Approve the administrator prompt once; it trusts the certificate and installs the app.

**Manual install:** `Orion-v3.1.msix` is signed with a self-signed certificate, so Windows must be told to trust it once per computer.

1. Download `Orion-Certificate.cer` and `Orion-v3.1.msix` from `Dist/1-Separate_Certificate_and_App`.
2. Right-click `Orion-Certificate.cer` → **Install Certificate** → **Local Machine** (administrator approval required) → **Place all certificates in the following store** → **Browse** → **Trusted People** → **Finish**.
3. Double-click `Orion-v3.1.msix` → **Install**. Orion then appears in the Start menu.
4. Only install the certificate if you received it directly from the developer. It is limited to the *Trusted People* store, which is enough for this package.

Future versions signed with the same certificate install without repeating step 2. Developers: rebuild the package with `build_exe/build_msix.ps1` (needs `certs/OrionSigning.pfx`, which is never committed).

## Time lock and permissions

Version 3.1 is time-locked and stops working after **31 December 2026**. After that date, or to request permission to use the app, please contact the developer, **Dr. L. Hulijeli**.

## Disclaimer

Provided as-is for geological interpretation support. Results must be verified by a qualified geologist before use in decisions.

&copy; 2026 Dr. L. Hulijeli. Use requires the developer's permission.

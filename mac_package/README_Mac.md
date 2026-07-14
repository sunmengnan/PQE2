# FAI IPQC Excel Extractor macOS Package

## What is included

This folder contains two macOS packaging options:

1. `FAI IPQC Excel Extractor.app`
   - A script-based macOS app bundle generated on Linux.
   - It can be double-clicked on macOS, but the Mac must have Python 3 with Tkinter available.
   - It installs `openpyxl` into the user Python environment if missing.

2. `build_mac_app.sh` + `FAI_IPQC_Excel_Extractor.spec`
   - PyInstaller build files for creating a real standalone macOS `.app` on a Mac.
   - The standalone app normally does not require the end user to install Python.

## Build a standalone app on Mac

1. Copy the whole `PQE2` project folder to a Mac.
2. Open the `mac_package` folder.
3. Right-click `build_mac_app.sh` and choose Open, or run it in Terminal:

```bash
./build_mac_app.sh
```

4. The result will be generated under:

- `mac_package/release/FAI IPQC Excel Extractor.app`
- `mac_package/release/FAI_IPQC_Excel_Extractor_macOS_Standalone.zip`

Send the ZIP file to users. Users only need to unzip it and double-click the app.

## Build ARM64 app through GitHub Actions

This project includes a GitHub Actions workflow:

- `.github/workflows/build-macos-arm64-app.yml`

After pushing the project to GitHub, the workflow runs on `macos-14` and verifies `uname -m = arm64`. It builds and uploads:

- `FAI_IPQC_Excel_Extractor_macOS_ARM64_Standalone.zip`

Manual build steps:

1. Open the GitHub repository.
2. Go to `Actions`.
3. Select `Build macOS ARM64 FAI IPQC Excel Extractor`.
4. Click `Run workflow`.
5. Download the artifact named `FAI_IPQC_Excel_Extractor_macOS_ARM64_Standalone`.

## Important limitation

A true standalone macOS binary must be built on macOS. Linux cannot directly cross-compile a real macOS PyInstaller app. The `.app` included directly in this folder is a launch-wrapper app for Macs that already have Python.

## First launch on macOS

If macOS says the app is from an unidentified developer:

1. Right-click the app.
2. Choose Open.
3. Confirm Open again.

This happens because the app is not Apple Developer ID signed or notarized.

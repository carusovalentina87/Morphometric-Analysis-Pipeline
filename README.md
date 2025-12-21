# Morphometric Analysis Pipeline for Archaeological Bitumen-based Samples

This repository provides a Python pipeline for **morphometric analysis of binary TIFF images**, designed for the quantitative characterization of pores and features selected from archaeological bitumen-based samples.

The pipeline extracts geometric properties from segmented features and produces:
- an **Excel workbook** with per-image statistics
- **TIFF figures** combining the binary image and histograms of morphometric properties

The workflow is lossless, reproducible, and intended for scientific use.

---

## Features

- Input restricted to **TIFF images** (`.tif`, `.tiff`)
- Automatic verification and binarization (0 / 255)
- External contour detection
- Morphometric measurements:
  - Area (mm²)
  - Major axis length (mm)
  - Minor axis length (mm)
  - Aspect ratio
- Per-image Excel sheets
- Per-image graphical summaries (image + histograms)
- Automatic creation of a dedicated `output/` folder

---

## Workflow

1. Load binary (or near-binary) TIFF images.
2. Verify binary values (0 / 255); apply thresholding if needed.
3. Detect external contours (features).
4. Compute morphometric properties in real units (mm).
5. Export results to:
   - an Excel workbook
   - TIFF figures with histograms

---

## Input Requirements

### Input folder (`--input`)
- Contains **binary or near-binary TIFF images**
- Foreground (features/pores): white (255)
- Background: black (0)

### Image format
- **Input:** `.tif`, `.tiff`
- **Output:** `.tiff`

---

## Output Structure

All outputs are saved inside a newly created `output/` folder:

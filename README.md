# Microsoft Word – Automated Image Resizing Based on Table Cell Dimensions

Automatically resize selected pictures inside Microsoft Word table cells while maintaining their original aspect ratio.

## Overview

This VBA project provides an automated way to resize inline pictures inside Microsoft Word tables.

The resizing procedure calculates the available space within the containing table cell, applies a 2-pixel margin, and adjusts the picture dimensions without distorting its original aspect ratio.

The implementation uses Microsoft Word's built-in **Normal.dotm** global template, allowing the functionality to be available across multiple Word documents without embedding macros in individual files.

## Repository Contents

| File                 | Description                                                                              |
| -------------------- | ---------------------------------------------------------------------------------------- |
| `autoresizefile.bas` | VBA standard module containing the image-resizing procedure and initialization routines. |

**Important:** The `.bas` file contains the standard module only. Automatic detection of picture selection requires the corresponding Word application event-handling implementation. Importing the standard module alone does not establish those events.

## Features

- Automatically detects selected inline pictures when the event handler is configured.
- Checks whether the selected picture is inside a table.
- Calculates available space using the containing table cell's dimensions.
- Maintains the original aspect ratio to prevent image distortion.
- Applies a 2-pixel margin around pictures.
- Centers pictures horizontally and vertically inside table cells.
- Works across multiple Word documents on the same computer.

---

## Setup Instructions

### 1. Open the Normal Template in the VBA Editor

1. Open Microsoft Word.
2. Press `ALT + F11` to launch the Visual Basic for Applications (VBA) Editor.
3. In the Project Explorer, locate **Normal**.
4. Expand **Normal**.
5. Expand **Modules**.
6. Open **Module1**.

**Important:** This project uses the existing `Normal.dotm` global template. There is no need to create a separate `.dotm` or `.docm` document.

### 2. Import or Insert the VBA Module

There are two ways to add the code.

**Option A – Import the BAS file**

1. In the VBA Editor, select the Normal project.
2. Right-click the project.
3. Select **Import File**.
4. Locate and select `autoresizefile.bas`.

**Option B – Copy the source code**

1. Open `autoresizefile.bas` using a text editor.
2. Copy its contents.
3. Paste the code into a standard module under Normal.

If importing the file, Word may create a module using the filename `autoresizefile`. You do not need to create another Module1 if the imported module already contains the required procedures.

### 3. Configure the Word Application Event Handler

The standard module contains the resizing procedure, but automatic detection of selected pictures requires an additional event-handling implementation.

The event handler must:

- Monitor Word application selection changes.
- Detect when an inline picture is selected.
- Call `FitSelectedImageInTableCell` with the current selection.
- Initialize when Word starts.

The original implementation references:

`ThisDocument.InitializeEvents`

That procedure must exist in the appropriate event-handling class module and be configured to monitor the Word application. The reference alone does not implement the event handler.

### 4. Save the Normal Template

1. Press `CTRL + S` in the VBA Editor.
2. Close the VBA Editor.
3. If Microsoft Word asks whether you want to save changes to Normal.dotm, select **Save**.

### 5. Restart Microsoft Word

1. Completely close Microsoft Word.
2. Reopen Microsoft Word.
3. Ensure macros are permitted to run.

The startup initialization routine should initialize the event handler when Word loads.

### 6. Insert and Select Pictures

1. Open any Word document.
2. Navigate to a table.
3. Insert an inline picture using **Insert → Pictures**.
4. Select the inserted picture.

When the event handler is properly configured, the resizing procedure will automatically fit the picture inside its containing table cell.

---

## How It Works

The resizing procedure calculates the available dimensions inside the table cell after subtracting the specified margins.

It then:

1. Retrieves the selected inline picture.
2. Identifies its containing table cell.
3. Retrieves the cell's width and height.
4. Subtracts the 2-pixel margins from both dimensions.
5. Calculates the picture's original aspect ratio.
6. Determines whether width or height is the limiting dimension.
7. Calculates the appropriate picture dimensions.
8. Resizes the picture without distortion.
9. Centers the picture horizontally.
10. Centers the picture vertically.

## Notes and Limitations

- Designed for inline pictures inside Word tables.
- Floating pictures with text wrapping are not supported.
- The picture must be selected for the resizing procedure to execute.
- Pictures outside tables are not supported.
- The script does not crop pictures.
- A 1.5-point margin is applied on each side (approximately 2 pixels at 96 DPI).
- The containing table cell's dimensions determine the maximum available picture size.
- Automatic execution requires the additional event-handling implementation.

## Important: Not Portable

**This implementation is installed locally in Microsoft Word's Normal.dotm template.**

The VBA code is stored in the local global template rather than inside individual Word documents.

Consequently:

- Documents opened on the computer with the installed implementation can use the functionality.
- Opening the same documents on another computer will not provide the functionality unless the code and event handler are installed there as well.
- Sharing a `.docx` document does not transfer the VBA functionality.
- Reinstalling Microsoft Office or Windows may require reinstalling the VBA code.

To use the functionality on another computer, install the required VBA components in that computer's Normal.dotm template and configure its event initialization.

---

## Recommended Table Formatting Settings

For consistent picture placement and predictable resizing behavior, configure the Word table using the following settings.

### 1. Paragraph Alignment

- Select the paragraph inside the table cell.
- Set paragraph alignment to **Center**.

This ensures pictures are horizontally centered.

### 2. Table Properties → Options

- Open **Table Properties**.
- Click **Options**.
- Set all four cell margins to `0`:
  - Top: `0`
  - Bottom: `0`
  - Left: `0`
  - Right: `0`

### 3. Table Properties → Row

- Open **Table Properties**.
- Go to the **Row** tab.
- Enable **Specify height**.
- Set the desired row height.
- Change **Row height is** to **Exactly**.

This prevents the table row from expanding or changing its height when inserting or dragging and dropping pictures.

### 4. Table Properties → Cell

- Open **Table Properties**.
- Go to the **Cell** tab.
- Set Vertical alignment to **Center**.

This ensures pictures remain vertically centered inside their respective cells.

**Important:** Apply these formatting settings to the table before inserting pictures for consistent results.

---

## Development Note

AI-assisted development. The code was reviewed and tested by the author.

---

<p align="center">
  <a href="https://github.com/arrianmunsod">
    <strong>← BACK TO ARRIAN'S GITHUB PROFILE</strong>
  </a>
  <br>
  <sub>Explore more projects, scripts, and technical documentation.</sub>
</p>

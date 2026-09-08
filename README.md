# AI-Assisted Excel Photo Automation

A practical example of using AI as a development aid to automate a repetitive documentation process in Excel.

## Overview

This project demonstrates an Excel VBA automation that inserts standardized photographs into predefined locations within an Excel documentation template.

The workflow is:

**Standardized photo filenames → filename mapping → VBA automation → automatic photo placement → completed documentation**

The automation was created and refined with AI assistance. I remained responsible for defining the process requirements, reviewing the proposed logic, testing the implementation, and confirming that the resulting automation behaved as expected.

This project is therefore an example of **AI-assisted process automation**, rather than an AI application itself.

## Problem

A repetitive documentation process required photographs to be manually placed into specific sections of an Excel template.

The manual approach involved:

- Renaming or identifying photographs according to their required position.
- Locating the corresponding section in the Excel template.
- Inserting each photograph manually.
- Repeating the process for multiple photographs.
- Checking that photographs were placed correctly.

This created unnecessary manual effort and increased the possibility of placement errors.

## Solution

The VBA automation uses a simple and reproducible convention:

1. Each photograph is assigned a standardized filename.
2. The filename identifies the photograph's required position.
3. The VBA macro determines the corresponding destination in the worksheet.
4. The photograph is inserted automatically.
5. The image is resized proportionally and centered within its designated area.
6. Existing photographs created by the automation are removed before a new run, allowing the process to be repeated safely.

The installed photographs use IDs `1` through `6`.

The dismantled photographs use IDs `1B` through `6B`.

## Photo Mapping

| ID | File | Type | Destination |
|---|---|---|---|
| 1 | `1.png` | Installed | `C8:X19` |
| 1B | `1B.png` | Dismantled | `AC8:AX19` |
| 2 | `2.png` | Installed | `C24:X35` |
| 2B | `2B.png` | Dismantled | `AC24:AX35` |
| 3 | `3.png` | Installed | `C40:X51` |
| 3B | `3B.png` | Dismantled | `AC40:AX51` |
| 4 | `4.png` | Installed | `C56:X67` |
| 4B | `4B.png` | Dismantled | `AC56:AX67` |
| 5 | `5.png` | Installed | `C72:X83` |
| 5B | `5B.png` | Dismantled | `AC72:AX83` |
| 6 | `6.png` | Installed | `C88:X99` |
| 6B | `6B.png` | Dismantled | `AC88:AX99` |

The mapping is also provided separately in `photo_mapping.csv`.

## Repository Structure

```text
AI-Assisted-Excel-Photo-Automation/
│
├── README.md
├── Documentation_Template.xlsx
├── InsertarFotos.bas
├── photo_mapping.csv
│
└── BF/
    ├── 1.png
    ├── 1B.png
    ├── 2.png
    ├── 2B.png
    ├── 3.png
    ├── 3B.png
    ├── 4.png
    ├── 4B.png
    ├── 5.png
    ├── 5B.png
    ├── 6.png
    └── 6B.png
```

## How to Reproduce

### 1. Prepare the files

Download or clone this repository and keep the folder structure unchanged.

The `BF` folder must remain in the same directory as `Documentation_Template.xlsx`.

### 2. Prepare the Excel workbook

Open `Documentation_Template.xlsx` in Excel.

Save a working copy as an Excel Macro-Enabled Workbook:

`Documentation.xlsm`

### 3. Import the VBA module

Open the VBA editor with:

`Alt + F11`

In the VBA editor:

1. Right-click the workbook project.
2. Select **Import File...**
3. Select `InsertPhotos.bas`.
4. Confirm that the module has been imported.

### 4. Verify the worksheet

The macro expects the worksheet to be named:

`Hoja4`

Do not rename this worksheet unless the VBA constant `WORKSHEET_NAME` is also changed.

### 5. Run the automation

Return to Excel and run:

`InsertPhotos`

The macro reads the photographs from the `BF` folder and places them into their corresponding destinations in the template.

### 6. Review the result

When the process finishes, the macro reports:

- Number of photographs inserted.
- Number of missing photographs, if any.
- Names of missing files.

The process can be run again. Photographs previously inserted by the automation are removed before the new run.

## Design Decisions

### Relative file path

The macro uses the workbook's own directory as the starting point:

```vb
photoFolder = ThisWorkbook.Path & Application.PathSeparator & PHOTO_FOLDER
```

This avoids dependency on a specific computer or user directory.

### Standardized naming

The filename is the primary identifier used by the automation.

For example:

`3.png`

is associated with the third installed-photo position, while:

`3B.png`

is associated with the corresponding dismantled-photo position.

This keeps the input convention simple and makes the process easy to reproduce.

### Proportional image handling

Photographs are inserted without distorting their original aspect ratio.

The macro calculates the available photo area, scales the image to fit, and centers it within that area.

### Repeatable execution

The automation removes photographs previously created by the macro before inserting a new set.

This makes the process repeatable without accumulating duplicate images.

## AI-Assisted Development

AI was used as a development aid during the creation and refinement of the VBA automation.

The collaboration included:

- Translating the manual process into explicit automation requirements.
- Developing and refining VBA logic.
- Identifying inconsistencies in the original implementation.
- Improving error handling and repeatability.
- Testing the macro against the Excel template.
- Reviewing the resulting output and correcting coordinate mappings.

The human role remained essential: the process requirements, expected behavior, testing, validation, and final decisions were reviewed manually.

## What This Demonstrates

This project demonstrates practical skills in:

- Process analysis and simplification.
- Workflow automation.
- Excel and VBA.
- File and naming conventions.
- Structured mapping between inputs and destinations.
- Error handling.
- Iterative testing and validation.
- Using AI as a development and problem-solving aid.
- Turning a repetitive manual task into a repeatable workflow.

## Limitations

This project intentionally keeps the scope simple.

It does not use:

- Machine learning models.
- Computer vision.
- An external AI API.
- A database.
- A web application.

The value of the project is in demonstrating how AI can assist with designing and refining a practical automation for a real repetitive process.

## License

This project is provided as a portfolio example for educational and demonstration purposes.

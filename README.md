# VBA Project - Color Palette Generator

## Description
This VBA project for Excel allows users to generate various shades of colors based on provided RGB values. Users enter RGB values in cells A3:C3 and then use the "Generate Color" button to create a color palette based on these values.

## Features
- **Dynamic Color Change**: Changing values in cells A3:C3 causes the background color of cell A1 to change to the corresponding RGB color.
- **Color Palette Generation**: After entering RGB values, users can generate a color palette by modifying the saturation and brightness by a specified percentage (1%, 2%, 5%, or 10%).

## Usage Instructions
1. **Entering RGB Values**: Enter R, G, and B values (from 0 to 255) in cells A3:C3. The color in cell A1 will change to reflect the chosen color.
2. **Generating Color Palette**: Click the "Generate Color" button. In response to the prompt, enter the desired accuracy for saturation and brightness changes (1%, 2%, 5%, or 10%). Upon acceptance, the script will generate a color palette on the sheet.

## Requirements
- Microsoft Excel with VBA macro support enabled.

## Installation
1. Open the Excel file "VBA Color Palette Maker" and go to the VBA editor (ALT + F11).
2. Double-click on Sheet2 (Color Generator) in the VBA editor and paste the code from `Code.txt`.
3. Once you have pasted the code, you can close the VBA editor window.
4. After pasting the code, you can immediately test the "Generate Color" button to activate the macro. Note that the pasted code will be operational right away in the current session.
5. To ensure that the macro remains functional in future sessions, save the file as an Excel Macro-Enabled Workbook (with the `.xlsm` extension). This step is crucial if you want the macro to be available the next time you open the file.
6. Remember to enable macros in your Excel settings to allow the macro to function correctly.

## Screenshots

![](https://github.com/Peciak95/VBA_Color_Maker/blob/main/Screenshots/1.png)
![](https://github.com/Peciak95/VBA_Color_Maker/blob/main/Screenshots/2.png)
![](https://github.com/Peciak95/VBA_Color_Maker/blob/main/Screenshots/3.png)
![](https://github.com/Peciak95/VBA_Color_Maker/blob/main/Screenshots/4.png)
![](https://github.com/Peciak95/VBA_Color_Maker/blob/main/Screenshots/5.png)

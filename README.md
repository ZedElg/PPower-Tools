# PPower Tools

A PowerPoint add in that adds a custom Ribbon with consulting style productivity tools, built in VBA.

<img width="1491" height="179" alt="Screenshot_PPower-Tools" src="https://github.com/user-attachments/assets/d4d08880-81f5-4a82-bb82-81c985ed0a73" />

---

## Features
The add in adds a custom Ribbon with tools for faster slide creation.

- Create chevrons, sticky notes, insert horizontal lines, harvey balls and convert shapes to text boxes
- Match shape height, width or full size, or resize shapes to fit text
- Select shapes with the same size or similar type
- Straighten lines and swap positions of two shapes
- Group shapes into rows or columns
- Split or merge text boxes, replace text, remove margins and insert footnotes
- Show a shortcut popup with helpful information

---

## Functions overview

| Group                | Command name                    | Macro name                    | Description                                                                 |
|----------------------|----------------------------------|--------------------------------|-----------------------------------------------------------------------------|
| Shapes               | Create Chevron Bar              | CreateChevronSequence          | Creates a sequence of chevrons. First one has a flat left side.             |
| Shapes             | Add Sticky                      | CreateSticky                   | Adds a sticky note rectangle at the first free slot on the slide.           |
| Shapes               | Add Horizontal Line             | AddHorizontalLine              | Inserts a horizontal line with the correct width based on slide layout.     |
| Shapes               | Convert Shapes to Text Boxes    | ConvertShapesToTextBoxes       | Replaces each selected shape with a text box containing its label text.     |
| Shapes        | Harvey Ball                     | HarveyBall                     | Inserts a Harvey Ball shape.                                                |
| Shapes        | Harvey Ball Toggle              | ToggleHarveyBall               | Toggles through Harvey Ball fill states.                                    |
| Shapes sizing        | Make Shapes Same Height         | MakeShapesSameHeight           | Makes all selected shapes match the first selected shape.                          |
| Shapes sizing        | Make Shapes Same Width          | MakeShapesSameWidth            | Makes all selected shapes match the first shape.                           |
| Shapes sizing        | Make Shapes Same Size           | MakeShapesSameSize             | Makes all selected shapes match both width and height of the first selected. |
| Shapes sizing        | Resize Shapes to Fit Text       | ResizeShapesToFitText          | Makes shapes height automatically shrink or expand to fit their text content.      |
| Shapes select        | Select Shapes with Same Size    | SelectShapesWithSameSize       | Selects all shapes on the slide that match the size of the first selected.  |
| Shapes select        | Select Similar Shapes           | SelectSimilarShapes            | Selects shapes similar in type to the first selected.                       |
| Shapes helpers       | Straighten Lines                | StraightenLines                | Straightens selected line shapes to perfect horizontal or vertical lines.   |
| Shapes helpers       | Swap Positions of Two Objects   | SwapPositionsOfTwoObjects      | Swaps position and size of exactly two selected shapes.                     |
| Text tools           | Split Text Box                  | SplitTextBoxes                 | Splits the selected text box into one box per paragraph while keeping formatting. |
| Text tools           | Merge Text Boxes                | MergeTextBoxes                 | Merges multiple text boxes into one. Uses formatting from the first selected. |
| Text tools           | Replace Text in Selected Shapes | ReplaceTextInSelectedShapes     | Finds and replaces text in all selected shapes with ...                    |
| Text tools                | Insert Footnote                 | InsertFootnote                 | Adds a footnote text box with layout aligned to slide bottom.               |
| Text tools           | Set Text Box Margins to Zero    | SetTextBoxMarginsToZero        | Removes inside margins from selected text boxes.                            |
| Align and arrange    | Align Shapes Center and Middle  | AlignShapesToCenterAndMiddle   | Centers selected shapes both vertically and horizontally.                   |
| Sorting and layout   | Group Shapes into Columns       | GroupByColumns                 | Groups selected shapes into multiple columns.                           |
| Sorting and layout   | Group Shapes into Rows          | GroupByRows                    | Groups selected shapes into rows.                                       |
| Other            | Show Shortcut Popup             | ShowShortcutPopup              | Displays a custom popup dialog with shortcuts or info.                      |

---

## Installation

1. Download the latest `PPower-Tools.ppam` from the `dist` folder.
2. In PowerPoint, go to  
   **File → Options → Add Ins → Manage: PowerPoint Add-Ins → Go…**
3. Click **Add New…**, select `PowerPointShortcutTools.ppam`, then enable it.
4. You should now see the new Ribbon tab called **PPower-Tools**.

---

## Development
If you have ideas for new features or want to report an issue, feel free to contact me through GitHub or open an issue in this repository.

This project is written in VBA and packaged as a `.ppam` add-in. 

The source code lives in the `src` folder:

- `src/modules` contains standard code modules  
- `src/classes` contains class modules  
- `src/forms` contains user forms for settings or dialogs  
- `src/ribbon/customUI.xml` defines the Ribbon layout and callbacks

To modify or extend the add in:

1. Open PowerPoint.
2. Load the `PPower-Tools_Demo.pptm` that contains the code in the example folder.
3. Press `ALT + F11` to open the VBA editor.
4. Import or edit the modules.


---

## License

This project is licensed under the [MIT License](LICENSE).

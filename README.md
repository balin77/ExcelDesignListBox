# ExcelDesignListBox
Excel Vba Listbox for Userforms with customizable .Font

The DesignListBox allows you to create a fully customizable ListBox in Excel. You may change the colors, make it bold,
change the FontStyle or FontSize - basically you can modify every entry like a Label. 

## How it works

  On runtime, the clsDesignListBox class creates a Frame on a Userform and in that it stores every value in a Label. 

### How to initialize it

  Add the two Class Files - clsDesignListBox and clsDesignListBoxObject - to your VBA project. 

  #### Create a Userform and define a new variable in its Event Userform_Initialize:

      Dim DesignListBox as new clsDesignListBox

  #### Create the DesignListBox with the .create procedure:
  (Params: Userform, Top, Left, Height, Width, ArrayWithData)
 
  IMPORTANT: 1D ARRAYS ARE NOT SUPPORTED YET

      DesignListBox.Create Me, 6, 6, 330, 534, InputArr
  
  #### Change the FontStyle of every Label as an example:
  
      Dim Labl
      For Each Labl In pDesignListBox.AllLabels
          Labl.Font.Name = "Arial Black"
      Next Labl
  
  #### Store DesingListBox in a private Variable defined at the top of the Userform Code
  
  ##### At top:
  
      Public WithEvents pDesignListBox As clsDesignListBox
  
  ##### Just after creating it:
  
      Set pDesignListBox = DesignListBox
  
  #### Use Events of the DesignListBox like this:
  
      Private sub pDesignListBox_Click()

      MsgBox pDesignListBox.SelectedValue

      end sub
  
### Procedures

|  Procedure |  Effect | Inputs |
| ------------ | ------------ | ------------ |
|  .Create    | Creates Designlistbox  | ParentUserForm as UserForm, Top as Long, Left as Long, Height as Long, Width as Long, ArrayWithData | 
|   .Clear   |   Clears all data in the Listbox.  |  |
|   .Fill     |   Provides new data. Old data will be deleted  | ArrayWithData  |
|   .Sort     |   Sorts data by a column. | ColumnNumber as Long (Zero Based), optional Descending as Boolean |
|   .Clear   |   Clears all data in the Listbox. The Frame is unaffected |  |
|   .SelectRow   |  Selects a certain row. Same result as changeing .ListIndex |  RowNumber as Long (Zero Based) |
|   .RaiseEventBevoreClick  |  Ignore |  |
|   .RaiseEventClick  |  Ignore |  |

    
### Properties

|  Property  |  Effect | Inputs | Output | Read / Write |
| ------------ | ------------ | ------------ | ------------ | ------------ |
|  .IsEmpty | Returns Boolean if the DesignListBox is empty  | | IsEmpty as Boolean | Read |
|  .ColumnsCount | Returns Long of the amount of Columns  | | ColumnsCount as Long | Read |
|  .RowsCount | Returns Long of the amount of Rows  | | RowsCount as Long | Read |
|  .RowHeight | Read / set Height as Long of each Row  | InpHeight as Long | RowHeight as Long | Read  / Write|
|  .ColumnWidths | Read / set Width as String for each Column | ColumnWidths as String (Notation: "15;20;30;") | ColumnWidths as String | Read  / Write |
|  .Headers | Read / set Headers as Boolean | IsOn as Boolean | IsOn as Boolean | Read  / Write |
|  .AllLabels | Returns all entries as Labels in a Collection |  | AllLabels as Collection | Read  |
|  .RowLabels | Returns all entries of a Row in a Collection | RowNumber as Long (Zero Based), InludingHeaders as Boolean (If you want HeaderLabels included)  | RowLabels as Collection | Read  |
|  .ColumnLabels | Returns all entries of a Column in a Collection | ColumnNumber as Long (Zero Based), InludingHeaders as Boolean (If you want HeaderLabels included)  | ColumnLabels as Collection | Read  |
|  .ExactLabel | Returns an exact entry as Object | ColumnNumber as Long (Zero Based), InludingHeaders as Boolean (If you want HeaderLabels included)  | ExactLabel as Object | Read  |
|  .HeadersLabels | Returns all HeaderLabels in a Collection |  | HeadersLabels as Collection | Read  |
|  .ColumnSource | Read / set ColumnNumber as Long of Column that returns .SelectedValue. | ColumnNumber as Long = 0 (Zero Based) | ColumnSource as Long | Read  / Write | 
|  .SelectedValue | Returns the value of the selected Row and the Column defined at .ColumnSource. The returned value is looked up in the InputArray provided at .Create or .Fill. Manual changes of the Label are ignored. |  | SelectedValue as Variant | Read | 
|  .TrueSelectedValue | Returns the value of the actual Label of the selected Row and the Column defined at .ColumnSource. Columns with ColumnWidth = 0 are ignored. |  | TrueSelectedValue as Variant | Read | 
|  .SelectionColor | Read / set SelectionColor as Long | ColorNumber as Long | SelectionColor as Long | Read / Write | 
|  .ListIndex | Read / set ListIndex as Long. (Any Row may be selected with .selectRow too) | RowNumber As Long  (Zero Based) | ListIndex As Long  (Zero Based) | Read / Write | 
|  .FreezeRows | Read / set the Rows that should stay put when scrolling. | RowsFromTop as Long (not Zero Based) | FreezeRows As Long  (not Zero Based) | Read / Write | 
|  .FreezeColumns | Read / set the Columns that should stay put when scrolling. | ColumnFromLef as Long (not Zero Based) | FreezeColumns As Long  (not Zero Based) | Read / Write | 
  
### Events

|  Event  |  Effect | Inputs | Output | 
| ------------ | ------------ | ------------ | ------------ | 
|  DesignListBox_BeforeClick | Fires after user has clicked on the DesignListBox, but before the provoked code executes | |  | 
|  DesignListBox_Click     | Fires after user has clicked on the DesignListBox and after the provoked code has been executed | |  | 
|  DesignListBox_Change    | Fires after .Create or .Fill has been executed | |  | 



  
Well so long my friends. Enjoy the code and create an issue, if you have any problems or improvements

Thank You
 
Raphael Gubler
  

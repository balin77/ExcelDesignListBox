# ExcelDesignListBox
Excel Vba Listbox for Userforms with customizable .Font

The DesignListBox allows you to create a fully customizable ListBox in Excel. You may change the colors, make it bold,
change the FontStyle or FontSize - basically you can modify every entry like a Label. 

How it works: 

  On runtime, the clsDesignListBox class creates a Frame on a Userform and in that it stores every value in a Label. 

How to initialize it:

  Add the two Class Files - clsDesignListBox and clsDesignListBoxObject - to your VBA project. 

  Create a Userform and define a new variable in its Event Userform_Initialize:

  Dim DesignListBox as new clsDesignListBox

  Create the DesignListBox with the .create procedure:
  (Params: Userform, Top, Left, Height, Width, ArrayWithData)
  IMPORTANT: 1D ARRAYS ARE NOT SUPPORTED YET

  DesignListBox.Create Me, 6, 6, 330, 534, InputArr
  
  Change the FontStyle of every Label as an example:
  
  Dim Labl
  For Each Labl In pDesignListBox.AllLabels
      Labl.Font.Name = "Arial Black"
  Next Labl
  
  Store DesingListBox in a private Variable defined at the top of the Userform Code
  
  At top:
  Public WithEvents pDesignListBox As clsDesignListBox
  
  Just after creating it:
  Set pDesignListBox = DesignListBox
  
  Use Events of the DesignListBox like this:
  
  Private sub pDesignListBox_Click()
  
  MsgBox pDesignListBox.SelectedValue
 
  end sub
  
Procedures:

|  Procedure |  Effect | Inputs |
| ------------ | ------------ | ------------ |
|  .Create    | Creates Designlistbox  | ParentUserForm as UserForm, Top as Long, Left as Long, Height as Long, Width as Long, ArrayWithData | 
|   .Clear   |   Clears all data in the Listbox.  |  |
|   .Fill     |   Provides new data. Old data will be deleted  | ArrayWithData  |
|   .Sort     |   Sorts data by a column. | ColumnNumber as Long (Zero Based), optional Descending as Boolean |
|   .Clear   |   Clears all data in the Listbox. The Frame is unaffected |  |
|   .SelectRow   |  Selects a certain row. Same result as changeing .ListIndex |  RowNumber as Long (Zero Based) |
|    |   |   |  
|   .RaiseEventBevoreClick  |  Ignore |  |
|   .RaiseEventClick  |  Ignore |  |

    
Properties:
  
  .IsEmpty                'returns Boolean if the DesignListBox is empty (Read) 
  .ColumnsCount           'returns Long of the amount of Columns (Read)
  .RowsCount              'returns Long of the amount of Rows (Read)
  .RowHeight              'read / set Height as Long of each Row (Read, write)
  .ColumnWidths           'read / set Width as String of each Column. (ColumnWidths Notation: "15;20;30;") (Read, write)
  .Headers                'read / set Headers as Boolean. (Read, write)
  .AllLabels              'returns all entries as Labels in a collection (Read)
  .RowLabels              'returns all entries of a Row in a collection. (RowNumber as Zero based Long, InludingHeaders as boolean if you want HeaderLabels included) (Read)
  .ColumnLabels           'returns all entries of a Column in a collection. (ColumnNumber as Zero based Long, InludingHeaders as boolean if you want HeaderLabels included) (Read)
  .ExactLabel             'returns an exact entry as object. (ColumnNumber as Zero based Long, RowNumber as Zero based Long) (Read)
  .HeadersLabels          'returns all HeaderLabels in a collection. (Read)
  .ColumnSource           'read / set ColumnNumber as Long of Column that should return .SelectedValue. Default = 0 (ColumnNumber as Zero based Long) (Read, write)
  .SelectedValue          'returns the value of the selected row and the Column defined at .ColumnSource. The returned Value is looked up in the InputArray. Manual changes of the Label are ignored. (Read)
  .TrueSelectedValue      'returns the value of the actual Label at the selected row and the Column defined at .ColumnSource. Columns with ColumnWidth = 0 are ignored. (Read)
  .SelectionColor         'read / set SelectionColor as Long (Read, write)
  .ListIndex              'read / set ListIndex as Long. (You may change the value also with .selectRow) (RowNumber as Zero based Long) (Read, write)
  .FreezeRows             'read / set the Rows that should stay put when scrolling. (RowsFromTop as Long (not Zero Based))(Read, write)
  .FreezeColumns          'read / set the Columns that should stay put when scrolling. (ColumnFromLef as Long (not Zero Based))(Read, write)
  
Events:
  
  _BevoreClick            'Fires bevore Click has been executed
  _Click                  'Fires after Click has been executed
  _Change                 'Fires after .Create or .Fill has been executed
  
Well so long my friends. Enjoy the code and create an issue, if you have any problems or improvements

Thank You
  
  

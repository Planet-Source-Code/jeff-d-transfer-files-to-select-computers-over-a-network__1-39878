Attribute VB_Name = "modComboBoxLoads"
Option Explicit

'THIS SUB FILLS THE COMBO BOX AS frmWSList OPENS
Public Sub GroupComboLoad(ctlList As Control)
   Dim oRS As Recordset
   Dim strSQL As String
   
   ctlList.Clear  'Clear the Combo Box
   
   'Build SQL Statement

      strSQL = "Select * from tblGroup"
   
   'Open Shapshot type Recordset
   Set oRS = dbXFER.OpenRecordset(strSQL, dbOpenSnapshot)
   
   'Loop through all row and load combo box)
   Do Until oRS.EOF
      ctlList.AddItem oRS!GroupName
      ctlList.ItemData(ctlList.NewIndex) = (oRS!GroupID)
      oRS.MoveNext
   Loop
   'Close the Snapshot Recordset
   oRS.Close
End Sub

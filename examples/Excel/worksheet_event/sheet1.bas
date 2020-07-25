option explicit

private sub worksheet_selectionChange(byVal target as range) ' {

   activeSheet.cells(2, 2).value = "You clicked " & target.address

end sub ' }

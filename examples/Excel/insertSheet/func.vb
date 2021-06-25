option explicit

sub main() ' {
  '
  ' Use "global variables" shOne and shTwo to
  ' refer to the corresponding sheets.
  '
    shOne.cells(2,2) = "This value was inserted by main()"
    shTwo.cells(2,2) = "The name of shOne is " & shOne.name

end sub ' }
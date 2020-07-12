option explicit

sub main(projectRootDir as variant) ' {
  '
  ' Note: the parameter(s) need to be declared as variant
  '       in Excel (apparently not so in Access).
  '
    activeSheet.cells(1,1) = "projectRootDir = " & projectRootDir

end sub ' }

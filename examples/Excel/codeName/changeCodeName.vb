option explicit

sub createSheet()

    dim sh as worksheet
    set sh = thisWorkbook.sheets.add
    sh.name = "Added in VBA"

  '
  ' https://stackoverflow.com/a/67904097/180275
  '
    on error resume next
    dim dummy as string
    dummy = thisWorkbook.VBProject.vbComponents(sh.CodeName).properties("codename")
    on error goto 0

    thisWorkbook.vbProject.vbComponents(sh.codeName).properties("_codeName") = "shXYZ"

end sub

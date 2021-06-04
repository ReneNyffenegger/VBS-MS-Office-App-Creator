'
' Provide the functionality to create Excel, Access or Word applications
' from the command line.
'
' The functions in this file should be called from a *.wsf file.
'
' Version 0.07
'
' See also https://renenyffenegger.ch/notes/Microsoft/Office/VBScript-App-Creator/
'

option explicit

dim fso
set fso = createObject("scripting.FileSystemObject")

function createOfficeApp(prod, fileName) ' {
 '
 '  Note: when creating an Access database, this function
 '  returns an application object. When creating an excel
 '  Worksheet, it returns an Excel Worksheet.
 '

    if fso.fileExists(fileName) then ' {
       on error resume next
       fso.deleteFile(fileName)

       if err.number = 70 then ' Permission denied {
          wscript.echo fileName & " could not be deleted (probably because it is in use)"
          set createOfficeApp = nothing
          exit function
       end if ' }
       on error goto 0

       wscript.echo fileName & " was deleted because it already existed"
    end if ' }

    dim fileSuffix, fileFormat
    fileSuffix = right(fileName, 5)

    dim app

    if     prod = "access" then ' {

           set createOfficeApp = createObject("access.application")
           createOfficeApp.newCurrentDatabase fileName, 0 ' 0: acNewDatabaseFormatUserDefault

           set app = createOfficeApp
    ' }
    elseIf prod = "excel"  then ' {

           set app             = createObject("excel.application")
         '
         ' createOfficeApp becomes a worksheet here, really...
         '
           set createOfficeApp = app.workBooks.add

         '
         ' Determine file format value based on extension of filename
         '
           if     fileSuffix = ".xlsb" then ' {
                  fileFormat = 50    ' xlExcel12

           elseif fileSuffix = ".xlsm" then
                  fileFormat = 52    ' xlOpenXMLWorkbookMacroEnabled

           else
                  wscript.echo fileName & " has  suffix that is not (yet?) supported"
                  set createOfficeApp = nothing
                  exit function
           end if ' }

           createOfficeApp.saveAs fileName, fileFormat
    ' }
    elseIf prod = "word"   then ' {
           set app             = createObject("word.application")

           set createOfficeApp = app.documents.add

         '
         ' Determine file format value based on extension of filename
         '
           if     fileSuffix = ".docm" then ' {
'                 fileFormat = 20    ' wdFormatFlatXMLMacroEnabled
                  fileFormat = 13    ' wdFormatXMLDocumentMacroEnabled

           elseif fileSuffix = ".dotm" then
'                 fileFormat = 22    ' wdFormatFlatXMLTemplateMacroEnabled
                  fileFormat = 15    ' wdFormatXMLTemplateMacroEnabled
           else

                  wscript.echo fileName & " has  suffix that is not (yet?) supported"
                  set createOfficeApp = nothing
                  exit function
           end if ' }


        '
        '  Note: saveAs2, not saveAs.
        '
           createOfficeApp.saveAs2 fileName, fileFormat

    end if ' }

    app.visible     = true

    if prod <> "word" then ' {
  '
  ' Keep application opened after scripts terminates
  '   https://stackoverflow.com/q/36282024/180275
  '
  ' In Word, userControl is read only and set to true if
  ' the application was created with createObject(), getObject() or opened
  ' with open()
  '
      app.userControl = true
    end if ' }

  '
  ' Add (type lib) reference to "Microsoft Visual Basic for Applications Extensibility 5.3"
  '
  '      2020-07-13: TODO: is this reference always present in Word documents?
  '
    call addReference(app, "{0002E157-0000-0000-C000-000000000046}", 5, 3)

end function ' }

function openOfficeApp(prod, fileName) ' {

    dim app
    if prod = "excel" then ' {
       set app = createObject("excel.application")
       dim updateLinks : updateLinks = false
       set openOfficeApp = app.workBooks.open(fileName, updateLinks)
    else
       wscript.echo("Todo. implement " & prod & " for openOfficeApp")
       set openOfficeApp = nothing
       exit function
    end if ' }

end function ' }

sub insertModule(app, moduleFilePath, moduleName, moduleType) ' {
 '
 '  moduleType:
 '    1 = vbext_ct_StdModule
 '    2 = vbext_ct_ClassModule
 '
 '  See also https://renenyffenegger.ch/notes/development/languages/VBA/modules/Common/00_ModuleLoader
 '

    if not fso.fileExists(moduleFilePath) then ' {
       wscript.echo moduleFilePath & " does not exist!"
       wscript.quit
    end if ' }

    dim vb_editor ' as vbe
    dim vb_proj   ' as VBProject
    dim vb_comps  ' as VBComponents

    if app.name = "Microsoft Word" then
       set vb_proj   = app.activeDocument.vbProject
    else

       on error resume next
       set vb_editor = app.vbe

       if err.number <> 0 then ' {
          if err.number = 1004 then
             wscript.echo("Unable to get reference of app.vbe: probably because macros are disabled")
             wscript.quit(-1)
          end if
          wscript.echo("Unexpected error when trying to get app.vba: " & err.number & " - " & err.description)
          wscript.quit(-1)
       end if ' }

       on error goto 0

       set vb_proj   = vb_editor.activeVBProject
    end if

    set vb_comps  = vb_proj.vbComponents

  '
  ' Check if a module by the given name already exists.
  ' If so, remove it.
  '
  ' If no module with the name moduleName exists, by default
  ' vb_comps(moduleName) throws a 'VBAProject: Subscript out of range'
  ' error.
  ' We're going to let such an error escape by embedding the
  ' statement between the following two 'on error …' statements:
  '
    on error resume next
    set comp = vb_comps(moduleName)
    on error goto 0

    dim comp      ' as VBComponent
    dim mdl       ' as codeModule

    if not isEmpty(comp) then

       set mdl = comp.codeModule
       dim nofLines
       nofLines = mdl.countOfLines
       mdl.deleteLines 1, nofLines

    else

       set comp = vb_comps.add(moduleType)
       set mdl  = comp.codeModule

    end if

 '
 '  2021-06-04  0.07  An absolute path is required when
 '                    calling addFromFile()
 '
    mdl.addFromFile fso.getAbsolutePathName(moduleFilePath)
    comp.name = moduleName

    if app.name = "Microsoft Access" then
       app.doCmd.close 5, comp.name, 1 ' 5=acModule, 1=acSaveYes
    end if

end sub ' }

sub addReference(app, guid, major, minor) ' {
  '
  ' guid identfies a type lib. Thus, the guid should be found in the
  ' Registry under HKEY_CLASSES_ROOT\TypeLib\
  '
  ' Note: guid probably needs the opening and closing curly paranthesis.
  '
    dim ref
    for each ref in app.VBE.activeVbProject.references
        if ref.guid = guid then
           wscript.echo "guid " & guid & " (" & ref.description & ") was already added"
           exit sub
        end if
    next

    call app.VBE.activeVbProject.references.addFromGuid (guid, major, minor)
end sub ' }

function currentDir() ' {
     dim wshShell
     set wshShell = createObject("WScript.Shell")

     currentDir = wshShell.CurrentDirectory & "\"

end function ' }

sub replaceThisWorksheetModule(app, moduleFilePath) ' {
 '
 '  Set the content of an Excel's ThisWorksheet module
 '
 '  TODO 2020-07-25 / Version 0.06 - This sub could probably call
 '     insertModule app, moduleFilePath, "thisWorksheet", 1
 '
    if not fso.fileExists(moduleFilePath) then ' {
       wscript.echo moduleFilePath & " does not exist!"
       wscript.quit
    end if ' }

    dim mdl
    set mdl = app.vbe.activeVBProject.vbComponents.item(1).codeModule
    call mdl.addFromFile (moduleFilePath)

end sub ' }

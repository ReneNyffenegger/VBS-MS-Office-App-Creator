'
'     Initialize form
'
'     Used once, at creation time of Excel Worksheet.
'
option explicit

const xLeft     = 20
const xRight    = 70
const btnHeight = 18
const tbWidth   = 70

sub initWorkbook() ' {

    dim btn as excel.button
    set btn  = sheet1.buttons.add( left := 30, top := 20, width := 65, height := 22)

    btn.caption  = "Enter values"
    btn.onAction = "openForm"

    initForm

end sub

private sub addLabelAndTextbox(dsgn as msForms.userForm, name as string, labelCaption as string, textBoxValue as string, top as long) ' {

    dim lb  as msForms.label
    dim tb  as msForms.textBox

    set tb     = dsgn.controls.add("forms.TextBox.1")
    set lb     = dsgn.controls.add("forms.Label.1"  )

    lb.left    = xLeft
    tb.left    = xRight

    lb.top     = top
    tb.top     = top - 2

    lb.width   =  xRight- xLeft -1
    tb.width   =  tbWidth

    lb.height  =  18
    tb.height  = btnHeight

    lb.caption = labelCaption
    tb.name    = name
    tb.text    = textBoxValue
end sub ' }

public sub initForm() ' {

    dim frm  as vbComponent
    dim dsgn as msForms.userForm
    set frm  = activeWorkbook.VBProject.vbcomponents("frmEnterValues")
    set dsgn = frm.designer

    frm.properties("Caption"     ) ="Enter values"
    frm.properties("width"       ) = xRight + tbWidth+ 40
    frm.properties("height"      ) = 100 + btnHeight + 50
    frm.properties("borderStyle" ) =   1 ' frmBorderStyleSingle

    dsgn.BackColor   =  rgb(255, 245, 235)
    dsgn.BorderColor =  rgb(200,  20,  40)

    addLabelAndTextBox dsgn, "valOne"  ,  "Value 1", "foo", 20
    addLabelAndTextBox dsgn, "valTwo"  ,  "Value 2", "bar", 42
    addLabelAndTextBox dsgn, "valThree",  "Value 3", "baz", 64

    dim btn as msForms.commandButton

    set btn = dsgn.controls.add("forms.commandButton.1")
    btn.left    = xLeft
    btn.top     = 100
    btn.width   =  50
    btn.height  = btnHeight
    btn.caption = "Ok"
    btn.name    = "ok"

    set btn = dsgn.controls.add("forms.commandButton.1")
    btn.left    = xRight+ tbWidth - 50
    btn.top     = 100
    btn.width   =  50
    btn.height  = btnHeight
    btn.caption = "Cancel"
    btn.name    = "cancel"

end sub ' }

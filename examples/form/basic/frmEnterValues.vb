option explicit


private sub userForm_activate() ' {
    with me.valOne
       .setFocus
       .selStart  = 0
       .selLength = len(.text)
    end with
end sub ' }


private sub ok_click() ' {
    okClicked me.valOne, me.valTwo, me.valThree   
    unload frmEnterValues
end sub ' }


private sub cancel_click() ' {
    unload frmEnterValues
end sub ' }

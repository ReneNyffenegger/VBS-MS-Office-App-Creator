option explicit

sub openForm() ' {

    frmEnterValues.show

end sub ' }

sub okClicked(valOne as string, valTwo as string, valThree as string) ' {
 '
 '  Called from frmEnterValues when user clicks 'OK'
 '
    with activeSheet
        .cells(5,2) = "Val one"  : .cells(5,3) = valOne
        .cells(6,2) = "Val two"  : .cells(6,3) = valTwo
        .cells(7,2) = "Val three": .cells(7,3) = valThree
    end with

end sub ' }

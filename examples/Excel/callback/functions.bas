option explicit

sub main(callback as variant) ' {

    callback.message("Main was started")

    cells(1,1) = "Hello world"

    callback.message("Finished.")

end sub ' }

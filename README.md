I'm teaching myself Access, I've been using YouTube a few different Templates. I've purchase some and the others I've gotten from Microsoft after I load Office 365. I've used these templates to see how they work. Also to use them as a starting point and as I address each obstacle in designing and building the database. The database is a Inventory control. The inventory consisted of complete systems as well as piece Parts. My current problem I'm working, I have a text box named Location. This is the location in the warehouse the product is. The complete are group together with a reference designator of GRP-1, GRP-2 and so on. I have another text box FileLocation, not visitable where I want to use VBA code to place the path to the .pdf file of the same name as the group. my code follows:

Private Sub Form_Open(Cancel As Integer)

Dim MyFile As String
Dim strLocation As String
strLocation = Location
MyFile = "C:\Inventory\PDF\""strLocation" & ".pdf"
If MyFile <> "c:\Inventory\PDF\GRP-*.pdf" Then
     FileLocation = "C:\Inventory\PDF\CoverSheet.pdf"
Else
     FileLocation = MyFile
End If
End Sub

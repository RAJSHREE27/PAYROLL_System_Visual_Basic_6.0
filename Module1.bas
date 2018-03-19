Attribute VB_Name = "Module1"
Public Sub Userlogin()
If ((Text1.Text = "160501") And (Text2.Text = "emp1")) Or ((Text1.Text = "160502") And (Text2.Text = "emp2")) Then
DataEnvironment1.Search Form2.Text1.Text

With DataEnvironment1.rsSearch
Form3.Show
Unload Me


If DataEnvironment1.rsSearch.EOF Then
MsgBox "Record Of Employee Not Found"


Form3.Text17.Text = ""
Form3.Text18.Text = ""
Form3.Text1.Text = ""
Form3.Text2.Text = ""
Form3.Text19.Text = ""
Form3.Text3.Text = ""
Form3.Text4.Text = ""
Form3.Text5.Text = ""
Form3.Text6.Text = ""
Form3.Text7.Text = ""
Form3.Text13.Text = ""
Form3.Text8.Text = ""
Form3.Text9.Text = ""
Form3.Text10.Text = ""
Form3.Text11.Text = ""
Form3.Text12.Text = ""
Form3.Text14.Text = ""
Form3.Text15.Text = ""
Else
MsgBox "Record Found"

Form3.Text16.Text = Val(DataEnvironment1.rsSearch.Fields("UserID"))

Form3.Text17.Text = DataEnvironment1.rsSearch.Fields("Emp Name")
Form3.Text18.Text = DataEnvironment1.rsSearch.Fields("Dept")
Form3.Text1.Text = DataEnvironment1.rsSearch.Fields("Issue Date")
Form3.Text2.Text = DataEnvironment1.rsSearch.Fields("Designation")
Form3.Text19.Text = DataEnvironment1.rsSearch.Fields("Email ID")
Form3.Text3.Text = DataEnvironment1.rsSearch.Fields("Basic Sal")
Form3.Text4.Text = DataEnvironment1.rsSearch.Fields("HRA")
Form3.Text5.Text = DataEnvironment1.rsSearch.Fields("DA")
Form3.Text6.Text = DataEnvironment1.rsSearch.Fields("Medical")
Form3.Text7.Text = DataEnvironment1.rsSearch.Fields("Special Allowance")
Form3.Text13.Text = DataEnvironment1.rsSearch.Fields("Total Allowance")
Form3.Text8.Text = DataEnvironment1.rsSearch.Fields("Insurance")
Form3.Text9.Text = DataEnvironment1.rsSearch.Fields("Loan")
Form3.Text10.Text = DataEnvironment1.rsSearch.Fields("Tax")
Form3.Text11.Text = DataEnvironment1.rsSearch.Fields("Leave")
Form3.Text12.Text = DataEnvironment1.rsSearch.Fields("Special Deduction")
Form3.Text14.Text = DataEnvironment1.rsSearch.Fields("Total Deduction")
Form3.Text15.Text = DataEnvironment1.rsSearch.Fields("Total Salary")
End If



.Close
End With
Else
MsgBox "Invalid Credentials"
End If

End Sub

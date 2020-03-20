import win32com.client #Make sure you have the right client 

path = "C:\ThePath\OfYourFolder\WithYourDocuments\\" 
# note the \\ at the end of the path name to prevent a SyntaxError

#Create the Application word
Application=win32com.client.gencache.EnsureDispatch("Word.Application")

# Compare documents
Application.CompareDocuments(Application.Documents.Open(path + "Original.docx"),
                             Application.Documents.Open(path + "Revised.docx"))

# Save the comparison document as "Comparison.docx"
Application.ActiveDocument.SaveAs (FileName = path + "Comparison.docx")
# Don't forget to quit your Application
Application.Quit()
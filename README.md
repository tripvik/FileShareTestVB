# Sample application to use ASP.NET impersonation.
Before running the application, update the **Form1.vb** file with the correct values of

    Dim userName As String = "<USERNAME>"   'Username Of the account To be impersonated
    Dim domain As String = "<DOMAIN>"       'Domain of the user account. Use localhost for local accounts
    Dim password As String = "<PASSWORD>"   'Account Password
         
    Dim src As String = "<SOURCE FILE PATH>"            'Path of the file to be copied
    Dim dest As String = "<DESTINATION FILE PATH>"      'Destination path. For the Azure fileshare use the network path instead of the drive letter
        
After making above changes, click on **Copy File** button. The file will be copied to the destination using the impersonated account.

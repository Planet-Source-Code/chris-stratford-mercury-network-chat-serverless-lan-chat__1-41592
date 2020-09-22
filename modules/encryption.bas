Attribute VB_Name = "Encryption_Module"
'Encryption Module
'Taken From PSCODE.COM
'No Authors Details Given
'Free Code, Adapted For Personal Use, I only have made minor changes


Public Function Decode(iString As String, iKEY As String) As String
    Dim Password As String
    Dim Words As String
    Dim Encrypted As String
    Dim Tempchar As String
    Dim Tempchar1 As String
    Dim Counter As Integer
    Dim TempAsc As Integer
    Dim TempAsc1 As Integer
    Counter = 1
    Password = iKEY
    Words = iString


    For X = 1 To Len(Words) 'loop for Each letter of the password
        
        Tempchar1 = Mid(Password, Counter, 1) 'get a Single letter of the password
        Tempchar = Mid(Words, X, 1) 'get a Single letter of the words
        
        TempAsc = Asc(Tempchar) 'convert the letter of the word To a number
        TempAsc1 = Asc(Tempchar1) 'convert the letter of the password To a number
        TempAsc = TempAsc - TempAsc1 ' subtract the two values
        
        'check to see if the value if greater th
        '     en 245. if it is, subtract 245 from it.
        'this makes sure that we don't go past t
        '     he highest ascii value
        If TempAsc < 0 Then TempAsc = TempAsc + 245
        
        Tempchar = Chr(TempAsc) 'convert the number back To a character
        'add the character to the end of the enc
        '     rypted string
        Encrypted = Encrypted & Tempchar
        Counter = Counter + 1 'incriment the counter
        
        'check to see if the counter is > the
        '     length of the password
        'if it is, set the counter to 1
        If Counter > Len(Password) Then Counter = 1
        
    Next X
    'show the encoded text in the textbox
    Decode = Encrypted
    
End Function


Public Function Encode(iString As String, iKEY As String) As String
    Dim Password As String
    Dim Words As String
    Dim Encrypted As String
    Dim Counter As Integer
    Dim Tempchar As String
    Counter = 1
    Password = iKEY
    Words = iString


    For X = 1 To Len(Words) 'loop for Each letter of the password
        
        Tempchar1 = Mid(Password, Counter, 1) 'get a Single letter of the password
        Tempchar = Mid(Words, X, 1) 'get a Single letter of the words
        
        TempAsc = Asc(Tempchar) 'convert the letter of the password To a number
        TempAsc1 = Asc(Tempchar1) 'convert the letter of the word To a number
        TempAsc = TempAsc + TempAsc1 ' add the two values
        
        'check to see if the value if greater than 245. if it is,
        'subtract 245 from it.
        'this makes sure that we don't go past the highest ascii value
        
        If TempAsc > 245 Then TempAsc = TempAsc - 245
        
        Tempchar = Chr(TempAsc) 'convert the number back To a character
        
        Encrypted = Encrypted & Tempchar 'add the character To the End of the encrypted String
        Counter = Counter + 1 'incriment the counter
        
        'check to see if the counter is > the
        '     length of the password
        'if it is, set the counter to 1
        If Counter > Len(Password) Then Counter = 1
        
    Next X
    'show the encoded text in the textbox
    Encode = Encrypted
    
End Function


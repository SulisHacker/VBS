REM Native VBS CRC32 library 1.0
REM Authentic work by SulisH@cker
REM SulisHacker@gmail.com © 2023 
REM All Rights Reserved
REM Any use of this code without the author's knowledge is prohibited!

Option Explicit

Class Crc32
    Private Table(255)
    'Sub Class_Initialize() 'Alternativnì: Public Sub Class_Initialize()
    Private pTimeElapsedLastCalc
    Private pCalculationSpeed
    
    Public Property Get TimeElapsedLastCalc
      TimeElapsedLastCalc = pTimeElapsedLastCalc
    End Property
    
    Public Property Get CalculationSpeed
      CalculationSpeed = pCalculationSpeed
    End Property   
	
	Private Function SHR8(ByRef Cislo) 'ALL THE MAGIC: Zde je hlavní kouzlo celého algoritmu, kde se obchází limitace absence operace bitového posuvu VBS a dále omezení VBS pøi dìlení 32Bitových èísel v horním rozsahu..
		SHR8 = (Cislo AND &H00FFFF00) \ 256 Or (Cislo AND &HFF000000) \ 256 AND &H00FFFFFF
	End Function
	
	Private Function SHR1(ByRef Cislo)
		SHR1 = (Cislo AND &H7FFFFFFF) \ 2 Or (Cislo AND &H80000000) \ 2 AND &H40000000
	End Function
    
    Private Sub Class_Initialize'()
    	pTimeElapsedLastCalc 			= 0
    	pCalculationSpeed               = 0
        Dim Poly: Poly 					= CLng(&HEDB88320)
        Dim Temp: Temp 					= CLng(0)
		Dim TableLength: TableLength 	= UBound(Table) ' - 1 'ZDE POZOR: UBound vrací horní hranici pole, zatímco Table.Length vrací poèet prvkù v poli..
        Dim i, j
        For i=0 To TableLength 
            Temp = (i)
            For j = 8 To 1 Step -1
                If (Temp And &H00000001) = 1 Then
                    Temp = CLng(SHR1(Temp)) Xor CLng(Poly)
                Else
                    Temp = CLng(SHR1(Temp))
                End If
            Next
            Table(i) = Temp
        Next
    End Sub
	
	Private Function ComputeCalculationSpeed(ByRef pTimeElapsedLastCalc, ByRef Delka)
		If pTimeElapsedLastCalc > 0 Then 'Zde mùže reálnì dojít k dìlení nulou, pokud jsou na vstupu data mmalé velikosti..
		  ComputeCalculationSpeed = Delka / pTimeElapsedLastCalc
		Else
		  ComputeCalculationSpeed = Delka
		End If 
	End Function

    Private Function ComputeBinCRC32CheckSum(ByRef oStream)
    	Dim StartTime: StartTime = Timer
    	Dim StopTime
        Dim CRC: CRC = &HFFFFFFFF
        Dim Delka: Delka = oStream.Size
        Dim Index: Index = CLng(0)
		Dim i, Znak, BinPosun

        For i = 1 To Delka
            Znak = (AscB(oStream.Read(1)))
            Index = Znak XOR (CRC AND &H000000FF)            
            BinPosun = SHR8(CRC)
            CRC = BinPosun XOR Table(Index)
        Next
        ComputeBinCRC32CheckSum = Not CRC
        StopTime = Timer
        pTimeElapsedLastCalc	= StopTime - StartTime
		pCalculationSpeed 		= ComputeCalculationSpeed(pTimeElapsedLastCalc, Delka)
    End Function
    
    Private Function ComputeStrCRC32CheckSum(ByRef oStream)
        Dim StartTime: StartTime = Timer
    	Dim StopTime
        Dim CRC: CRC = &HFFFFFFFF
        Dim Delka: Delka = oStream.Size
        Dim Index, i, Znak, BinPosun
		oStream.Position = 0 'Nastavení pozice na zaèátek, aby bylo možné zaèít èíst obsažená data
        For i = 1 To Delka
            Znak  = AscB(oStream.ReadText(1))
            Index = Znak XOR (CRC AND &H000000FF)
            BinPosun = SHR8(CRC)            
            CRC = BinPosun XOR Table(Index)
        Next
        ComputeStrCRC32CheckSum = Not CRC
        StopTime				= Timer
        pTimeElapsedLastCalc	= StopTime - StartTime
		pCalculationSpeed 		= ComputeCalculationSpeed(pTimeElapsedLastCalc, Delka)
    End Function
    
    'https://www.w3schools.com/asp/ado_ref_stream.asp
	'https://www.devguru.com/content/technologies/ado/objects-stream.html
    Public Function ComputeFileCRC32CheckSum(ByRef FileName)
    	Dim oStream: Set oStream = CreateObject("ADODB.Stream")
		Const adTypeBinary = 1
    	oStream.Type = adTypeBinary
    	oStream.Open
    	oStream.LoadFromFile FileName
    		ComputeFileCRC32CheckSum = Hex(ComputeBinCRC32CheckSum(oStream))
    	oStream.Close
    	Set oStream = Nothing    
    End Function
    
    Public Function ComputeStringCRC32CheckSum(ByRef strData)
    	Dim oStream: Set oStream = CreateObject("ADODB.Stream")
    	oStream.Charset = "ASCII" 'Defaultnì se použije unicode které bude vkládat BOM byty, takže nebude sedìt vložená velikost dat..
    	oStream.Type = 2 'adTypeText
    	oStream.Mode = 16
    	oStream.Open
    	oStream.WriteText(strData)
    		ComputeStringCRC32CheckSum = Hex(ComputeStrCRC32CheckSum(oStream))
    	oStream.Close
    	Set oStream = Nothing    
    End Function
End Class

Dim CRC, strOut
  strOut = ""

Set CRC = New CRC32
  'strOut = strOut & "CRC32					= " 				& CRC.ComputeFileCRC32CheckSum("TEST_FILE.dll") & vbCrlf
  strOut = strOut &  "CRC32					= "					& (CRC.ComputeStringCRC32CheckSum("A")) & vbCrlf 'CRC32 of char "A" = D3D99E8B
  strOut = strOut &  "Time elapsed last CRC32 calculation 		= " & CRC.TimeElapsedLastCalc & "s"		& vbCrlf
  strOut = strOut &  "Speed of last CRC32 calculation		= " & CRC.CalculationSpeed & "B/s"	& vbCrlf
Set CRC = Nothing

WScript.Echo strOut
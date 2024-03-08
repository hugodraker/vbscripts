'LICENCE PUBLIC DOMAIN
'NO WARRANTY
'WriteWildFiles strUseFolder, ".p0" 'Write to every file *.p0* at offset values specified
'NO FILE FORMAT CHECKING WHATSOEVER, important data can be overwritten
Option Explicit
Public name(200), surname(200), birthdate(200), fn, ln, bd, mem, i, x, ByteArray, path
Dim strUseFolder, B, BinaryStream, d()
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
name(0) = "James"
name(1) = "Robert"
name(2) = "John"
name(3) = "Michael"
name(4) = "David"
name(5) = "William"
name(6) = "Richard"
name(7) = "Joseph"
name(8) = "Thomas"
name(9) = "Christopher"
name(10) = "Charles"
name(11) = "Daniel"
name(12) = "Matthew"
name(13) = "Anthony"
name(14) = "Mark"
name(15) = "Donald"
name(16) = "Steven"
name(17) = "Andrew"
name(18) = "Paul"
name(19) = "Joshua"
name(20) = "Kenneth"
name(21) = "Kevin"
name(22) = "Brian"
name(23) = "George"
name(24) = "Timothy"
name(25) = "Ronald"
name(26) = "Jason"
name(27) = "Edward"
name(28) = "Jeffrey"
name(29) = "Ryan"
name(30) = "Jacob"
name(31) = "Gary"
name(32) = "Nicholas"
name(33) = "Eric"
name(34) = "Jonathan"
name(35) = "Stephen"
name(36) = "Larry"
name(37) = "Justin"
name(38) = "Scott"
name(39) = "Brandon"
name(40) = "Benjamin"
name(41) = "Samuel"
name(42) = "Gregory"
name(43) = "Alexander"
name(44) = "Patrick"
name(45) = "Frank"
name(46) = "Raymond"
name(47) = "Jack"
name(48) = "Dennis"
name(49) = "Jerry"
name(50) = "Tyler"
name(51) = "Aaron"
name(52) = "Jose"
name(53) = "Adam"
name(54) = "Nathan"
name(55) = "Henry"
name(56) = "Zachary"
name(57) = "Douglas"
name(58) = "Peter"
name(59) = "Kyle"
name(60) = "Noah"
name(61) = "Ethan"
name(62) = "Jeremy"
name(63) = "Walter"
name(64) = "Christian"
name(65) = "Keith"
name(66) = "Roger"
name(67) = "Terry"
name(68) = "Austin"
name(69) = "Sean"
name(70) = "Gerald"
name(71) = "Carl"
name(72) = "Harold"
name(73) = "Dylan"
name(74) = "Arthur"
name(75) = "Lawrence"
name(76) = "Jordan"
name(77) = "Jesse"
name(78) = "Bryan"
name(79) = "Billy"
name(80) = "Bruce"
name(81) = "Gabriel"
name(82) = "Joe"
name(83) = "Logan"
name(84) = "Alan"
name(85) = "Juan"
name(86) = "Albert"
name(87) = "Willie"
name(88) = "Elijah"
name(89) = "Wayne"
name(90) = "Randy"
name(91) = "Vincent"
name(92) = "Mason"
name(93) = "Roy"
name(94) = "Ralph"
name(95) = "Bobby"
name(96) = "Russell"
name(97) = "Bradley"
name(98) = "Philip"
name(99) = "Eugene"
name(100) = "Mary"
name(101) = "Patricia"
name(102) = "Jennifer"
name(103) = "Linda"
name(104) = "Elizabeth"
name(105) = "Barbara"
name(106) = "Susan"
name(107) = "Jessica"
name(108) = "Sarah"
name(109) = "Karen"
name(110) = "Lisa"
name(111) = "Nancy"
name(112) = "Betty"
name(113) = "Sandra"
name(114) = "Margaret"
name(115) = "Ashley"
name(116) = "Kimberly"
name(117) = "Emily"
name(118) = "Donna"
name(119) = "Michelle"
name(120) = "Carol"
name(121) = "Amanda"
name(122) = "Melissa"
name(123) = "Deborah"
name(124) = "Stephanie"
name(125) = "Dorothy"
name(126) = "Rebecca"
name(127) = "Sharon"
name(128) = "Laura"
name(129) = "Cynthia"
name(130) = "Amy"
name(131) = "Kathleen"
name(132) = "Angela"
name(133) = "Shirley"
name(134) = "Brenda"
name(135) = "Emma"
name(136) = "Anna"
name(137) = "Pamela"
name(138) = "Nicole"
name(139) = "Samantha"
name(140) = "Katherine"
name(141) = "Christine"
name(142) = "Helen"
name(143) = "Debra"
name(144) = "Rachel"
name(145) = "Carolyn"
name(146) = "Janet"
name(147) = "Maria"
name(148) = "Catherine"
name(149) = "Heather"
name(150) = "Diane"
name(151) = "Olivia"
name(152) = "Julie"
name(153) = "Joyce"
name(154) = "Victoria"
name(155) = "Ruth"
name(156) = "Virginia"
name(157) = "Lauren"
name(158) = "Kelly"
name(159) = "Christina"
name(160) = "Joan"
name(161) = "Evelyn"
name(162) = "Judith"
name(163) = "Andrea"
name(164) = "Hannah"
name(165) = "Megan"
name(166) = "Cheryl"
name(167) = "Jacqueline"
name(168) = "Martha"
name(169) = "Madison"
name(170) = "Teresa"
name(171) = "Gloria"
name(172) = "Sara"
name(173) = "Janice"
name(174) = "Ann"
name(175) = "Kathryn"
name(176) = "Abigail"
name(177) = "Sophia"
name(178) = "Frances"
name(179) = "Jean"
name(180) = "Alice"
name(181) = "Judy"
name(182) = "Isabella"
name(183) = "Julia"
name(184) = "Grace"
name(185) = "Amber"
name(186) = "Denise"
name(187) = "Danielle"
name(188) = "Marilyn"
name(189) = "Beverly"
name(190) = "Charlotte"
name(191) = "Natalie"
name(192) = "Theresa"
name(193) = "Diana"
name(194) = "Brittany"
name(195) = "Doris"
name(196) = "Kayla"
name(197) = "Alexis"
name(198) = "Lori"
name(199) = "Marie"

surname(0) = "Smith"
surname(1) = "Johnson"
surname(2) = "Williams"
surname(3) = "Brown"
surname(4) = "Jones"
surname(5) = "Garcia"
surname(6) = "Miller"
surname(7) = "Davis"
surname(8) = "Rodriguez"
surname(9) = "Martinez"
surname(10) = "Hernandez"
surname(11) = "Lopez"
surname(12) = "Gonzalez"
surname(13) = "Wilson"
surname(14) = "Anderson"
surname(15) = "Thomas"
surname(16) = "Taylor"
surname(17) = "Moore"
surname(18) = "Jackson"
surname(19) = "Martin"
surname(20) = "Lee"
surname(21) = "Perez"
surname(22) = "Thompson"
surname(23) = "White"
surname(24) = "Harris"
surname(25) = "Sanchez"
surname(26) = "Clark"
surname(27) = "Ramirez"
surname(28) = "Lewis"
surname(29) = "Robinson"
surname(30) = "Walker"
surname(31) = "Young"
surname(32) = "Allen"
surname(33) = "King"
surname(34) = "Wright"
surname(35) = "Scott"
surname(36) = "Torres"
surname(37) = "Nguyen"
surname(38) = "Hill"
surname(39) = "Flores"
surname(40) = "Green"
surname(41) = "Adams"
surname(42) = "Nelson"
surname(43) = "Baker"
surname(44) = "Hall"
surname(45) = "Rivera"
surname(46) = "Campbell"
surname(47) = "Mitchell"
surname(48) = "Carter"
surname(49) = "Roberts"
surname(50) = "Gomez"
surname(51) = "Phillips"
surname(52) = "Evans"
surname(53) = "Turner"
surname(54) = "Diaz"
surname(55) = "Parker"
surname(56) = "Cruz"
surname(57) = "Edwards"
surname(58) = "Collins"
surname(59) = "Reyes"
surname(60) = "Stewart"
surname(61) = "Morris"
surname(62) = "Morales"
surname(63) = "Murphy"
surname(64) = "Cook"
surname(65) = "Rogers"
surname(66) = "Gutierrez"
surname(67) = "Ortiz"
surname(68) = "Morgan"
surname(69) = "Cooper"
surname(70) = "Peterson"
surname(71) = "Bailey"
surname(72) = "Reed"
surname(73) = "Kelly"
surname(74) = "Howard"
surname(75) = "Ramos"
surname(76) = "Kim"
surname(77) = "Cox"
surname(78) = "Ward"
surname(79) = "Richardson"
surname(80) = "Watson"
surname(81) = "Brooks"
surname(82) = "Chavez"
surname(83) = "Wood"
surname(84) = "James"
surname(85) = "Bennett"
surname(86) = "Gray"
surname(87) = "Mendoza"
surname(88) = "Ruiz"
surname(89) = "Hughes"
surname(90) = "Price"
surname(91) = "Alvarez"
surname(92) = "Castillo"
surname(93) = "Sanders"
surname(94) = "Patel"
surname(95) = "Myers"
surname(96) = "Long"
surname(97) = "Ross"
surname(98) = "Foster"
surname(99) = "Jimenez"
surname(100) = "Powell"
surname(101) = "Jenkins"
surname(102) = "Perry"
surname(103) = "Russell"
surname(104) = "Sullivan"
surname(105) = "Bell"
surname(106) = "Coleman"
surname(107) = "Butler"
surname(108) = "Henderson"
surname(109) = "Barnes"
surname(110) = "Gonzales"
surname(111) = "Fisher"
surname(112) = "Vasquez"
surname(113) = "Simmons"
surname(114) = "Romero"
surname(115) = "Jordan"
surname(116) = "Patterson"
surname(117) = "Alexander"
surname(118) = "Hamilton"
surname(119) = "Graham"
surname(120) = "Reynolds"
surname(121) = "Griffin"
surname(122) = "Wallace"
surname(123) = "Moreno"
surname(124) = "West"
surname(125) = "Cole"
surname(126) = "Hayes"
surname(127) = "Bryant"
surname(128) = "Herrera"
surname(129) = "Gibson"
surname(130) = "Ellis"
surname(131) = "Tran"
surname(132) = "Medina"
surname(133) = "Aguilar"
surname(134) = "Stevens"
surname(135) = "Murray"
surname(136) = "Ford"
surname(137) = "Castro"
surname(138) = "Marshall"
surname(139) = "Owens"
surname(140) = "Harrison"
surname(141) = "Fernandez"
surname(142) = "Mcdonald"
surname(143) = "Woods"
surname(144) = "Washington"
surname(145) = "Kennedy"
surname(146) = "Wells"
surname(147) = "Vargas"
surname(148) = "Henry"
surname(149) = "Chen"
surname(150) = "Freeman"
surname(151) = "Webb"
surname(152) = "Tucker"
surname(153) = "Guzman"
surname(154) = "Burns"
surname(155) = "Crawford"
surname(156) = "Olson"
surname(157) = "Simpson"
surname(158) = "Porter"
surname(159) = "Hunter"
surname(160) = "Gordon"
surname(161) = "Mendez"
surname(162) = "Silva"
surname(163) = "Shaw"
surname(164) = "Snyder"
surname(165) = "Mason"
surname(166) = "Dixon"
surname(167) = "Munoz"
surname(168) = "Hunt"
surname(169) = "Hicks"
surname(170) = "Holmes"
surname(171) = "Palmer"
surname(172) = "Wagner"
surname(173) = "Black"
surname(174) = "Robertson"
surname(175) = "Boyd"
surname(176) = "Rose"
surname(177) = "Stone"
surname(178) = "Salazar"
surname(179) = "Fox"
surname(180) = "Warren"
surname(181) = "Mills"
surname(182) = "Meyer"
surname(183) = "Rice"
surname(184) = "Schmidt"
surname(185) = "Garza"
surname(186) = "Daniels"
surname(187) = "Ferguson"
surname(188) = "Nichols"
surname(189) = "Stephens"
surname(190) = "Soto"
surname(191) = "Weaver"
surname(192) = "Ryan"
surname(193) = "Gardner"
surname(194) = "Payne"
surname(195) = "Grant"
surname(196) = "Dunn"
surname(197) = "Kelley"
surname(198) = "Spencer"
surname(199) = "Hawkins"

birthdate(0) = "20060806"
birthdate(1) = "19530820"
birthdate(2) = "20070210"
birthdate(3) = "19880101"
birthdate(4) = "19510820"
birthdate(5) = "19631210"
birthdate(6) = "19820310"
birthdate(7) = "20100210"
birthdate(8) = "20100812"
birthdate(9) = "19550920"
birthdate(10) = "20110101"
birthdate(11) = "20100202"
birthdate(12) = "19780103"
birthdate(13) = "19990603"
birthdate(14) = "19540826"
birthdate(15) = "19730906"
birthdate(16) = "20070410"
birthdate(17) = "19981115"
birthdate(18) = "20100331"
birthdate(19) = "19890804"
birthdate(20) = "19470606"
birthdate(21) = "20020928"
birthdate(22) = "20080706"
birthdate(23) = "19910226"
birthdate(24) = "19960711"
birthdate(25) = "19950715"
birthdate(26) = "19710903"
birthdate(27) = "19890729"
birthdate(28) = "19560822"
birthdate(29) = "19930105"
birthdate(30) = "19470302"
birthdate(31) = "19631031"
birthdate(32) = "19480221"
birthdate(33) = "19510810"
birthdate(34) = "20001230"
birthdate(35) = "19920401"
birthdate(36) = "19660725"
birthdate(37) = "20090813"
birthdate(38) = "19470506"
birthdate(39) = "19741102"
birthdate(40) = "19701212"
birthdate(41) = "19970121"
birthdate(42) = "19990128"
birthdate(43) = "19570916"
birthdate(44) = "19780422"
birthdate(45) = "19750421"
birthdate(46) = "19881213"
birthdate(47) = "19930328"
birthdate(48) = "19960427"
birthdate(49) = "19631009"
birthdate(50) = "19910322"
birthdate(51) = "19890719"
birthdate(52) = "19560122"
birthdate(53) = "19530203"
birthdate(54) = "19781121"
birthdate(55) = "20100407"
birthdate(56) = "19680224"
birthdate(57) = "19841019"
birthdate(58) = "19600321"
birthdate(59) = "19960202"
birthdate(60) = "19620507"
birthdate(61) = "19790529"
birthdate(62) = "19920715"
birthdate(63) = "20050801"
birthdate(64) = "20100326"
birthdate(65) = "19820319"
birthdate(66) = "19540606"
birthdate(67) = "19550226"
birthdate(68) = "19620706"
birthdate(69) = "20020303"
birthdate(70) = "19620417"
birthdate(71) = "20000516"
birthdate(72) = "19610724"
birthdate(73) = "20080311"
birthdate(74) = "19681019"
birthdate(75) = "19580515"
birthdate(76) = "19620128"
birthdate(77) = "19861122"
birthdate(78) = "19770309"
birthdate(79) = "19681130"
birthdate(80) = "20010701"
birthdate(81) = "19841019"
birthdate(82) = "19820520"
birthdate(83) = "20070516"
birthdate(84) = "19640609"
birthdate(85) = "19960628"
birthdate(86) = "19960403"
birthdate(87) = "19701115"
birthdate(88) = "19830812"
birthdate(89) = "19500227"
birthdate(90) = "19480901"
birthdate(91) = "19810204"
birthdate(92) = "19971226"
birthdate(93) = "20080707"
birthdate(94) = "19531101"
birthdate(95) = "19830906"
birthdate(96) = "19761202"
birthdate(97) = "19451023"
birthdate(98) = "19671205"
birthdate(99) = "19560112"
birthdate(100) = "19990105"
birthdate(101) = "19660301"
birthdate(102) = "19801210"
birthdate(103) = "19560407"
birthdate(104) = "19851208"
birthdate(105) = "19621119"
birthdate(106) = "19890624"
birthdate(107) = "19911114"
birthdate(108) = "19951116"
birthdate(109) = "19750822"
birthdate(110) = "19500913"
birthdate(111) = "19600728"
birthdate(112) = "20070209"
birthdate(113) = "19550513"
birthdate(114) = "20010226"
birthdate(115) = "19810810"
birthdate(116) = "20120926"
birthdate(117) = "19500426"
birthdate(118) = "19750207"
birthdate(119) = "19520402"
birthdate(120) = "20100530"
birthdate(121) = "19450426"
birthdate(122) = "19970911"
birthdate(123) = "20000730"
birthdate(124) = "20040127"
birthdate(125) = "19500929"
birthdate(126) = "19720309"
birthdate(127) = "19620903"
birthdate(128) = "19990529"
birthdate(129) = "19740504"
birthdate(130) = "20061204"
birthdate(131) = "19570514"
birthdate(132) = "19621210"
birthdate(133) = "19541124"
birthdate(134) = "19540403"
birthdate(135) = "20040211"
birthdate(136) = "19840603"
birthdate(137) = "19820523"
birthdate(138) = "19541110"
birthdate(139) = "20030103"
birthdate(140) = "19870420"
birthdate(141) = "19681112"
birthdate(142) = "19791126"
birthdate(143) = "19720428"
birthdate(144) = "19500302"
birthdate(145) = "19610425"
birthdate(146) = "19530521"
birthdate(147) = "19570704"
birthdate(148) = "19610426"
birthdate(149) = "19730517"
birthdate(150) = "19480518"
birthdate(151) = "20060521"
birthdate(152) = "20090331"
birthdate(153) = "19780519"
birthdate(154) = "19780409"
birthdate(155) = "19671219"
birthdate(156) = "20060316"
birthdate(157) = "19700209"
birthdate(158) = "19520724"
birthdate(159) = "19980122"
birthdate(160) = "19710703"
birthdate(161) = "19610608"
birthdate(162) = "19720619"
birthdate(163) = "19510724"
birthdate(164) = "19531222"
birthdate(165) = "20090122"
birthdate(166) = "20100107"
birthdate(167) = "19840212"
birthdate(168) = "19490124"
birthdate(169) = "19601219"
birthdate(170) = "19690106"
birthdate(171) = "20001103"
birthdate(172) = "19460118"
birthdate(173) = "19471205"
birthdate(174) = "19560629"
birthdate(175) = "19890221"
birthdate(176) = "19941004"
birthdate(177) = "19890118"
birthdate(178) = "19750831"
birthdate(179) = "19820314"
birthdate(180) = "19650224"
birthdate(181) = "19950822"
birthdate(182) = "19571107"
birthdate(183) = "19910914"
birthdate(184) = "19570624"
birthdate(185) = "19700122"
birthdate(186) = "19870718"
birthdate(187) = "19980121"
birthdate(188) = "19500708"
birthdate(189) = "20080314"
birthdate(190) = "19971001"
birthdate(191) = "19780207"
birthdate(192) = "19740822"
birthdate(193) = "19750520"
birthdate(194) = "19651031"
birthdate(195) = "19790731"
birthdate(196) = "19790926"
birthdate(197) = "20000807"
birthdate(198) = "19990119"
birthdate(199) = "19881024" 'yyyymmdd =concatenate("birthdate(",B1,") = \"",I1,"\"")  =randbetween(date(1945,1,1),date(2012,12,31))

If WScript.Arguments.Count = 0 then 
    strUseFolder = ""
Else 
    strUseFolder = WScript.Arguments(0)
End if

WriteWildFiles strUseFolder, ".p0" 'Write to every file *.p0*

WScript.Quit

Public Function randomname(S)
Dim min, max
   if S="m" then 
      min=0
      max=99 
   elseif S="f" then 
      min=100
      max=199
   else 
      min=0
      max=199
   end if

   Randomize
   fn=(Int((max-min+1)*Rnd+min))
   max=199
   min=0
   Randomize
   ln=(Int((max-min+1)*Rnd+min))
   bd=(Int((max-min+1)*Rnd+min))
End Function

Function RandomDateYYYYMMDD(startDate, endDate)
    startDate = CDate(startDate)
    endDate = CDate(endDate)
    Call Randomize()
    RandomDateYYYYMMDD = DateAdd( _ 
        "d" _ 
        , Fix( DateDiff("d", startDate, endDate ) * Rnd ) _ 
        , startDate _ 
    )
    RandomDateYYYYMMDD= Year(RandomDateYYYYMMDD)&""&LPad(Month(RandomDateYYYYMMDD), "0", 2) &""&LPad(Day(RandomDateYYYYMMDD), "0", 2)
End Function

Function RandomDateMMDDYYYY(startDate, endDate)
    startDate = CDate(startDate)
    endDate = CDate(endDate)
    Call Randomize()
    RandomDateMMDDYYYY = DateAdd( _ 
        "d" _ 
        , Fix( DateDiff("d", startDate, endDate ) * Rnd ) _ 
        , startDate _ 
    )
    RandomDateMMDDYYYY = LPad(Month(RandomDateMMDDYYYY), "0", 2) &"-"&LPad(Day(RandomDateMMDDYYYY), "0", 2) &"-"&Year(RandomDateMMDDYYYY )
End Function

Private Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function

Private Function OverWritePersonName (strFilename)
    randomname("f")
    'WScript.Echo name(fn) & " " & surname(ln) & " - " & birthdate(bd)

    'WScript.Echo "Writing the file data"
    Set BinaryStream = CreateObject("ADODB.Stream")
    BinaryStream.Type = adTypeBinary
    BinaryStream.Open
    BinaryStream.LoadFromFile(strFilename)
    BinaryStream.Position = &H3E 'SET OFFSET

    Set mem = CreateObject("System.IO.MemoryStream")
    mem.SetLength(0)

    i=ucase(name(fn)) 
    ReDim d(&H34) 'd(Len(i)*2)
    For x = 1 To Len(i) 
        d((x-1)*2) = asc(Mid(i,x,1))
    Next
    For Each B in d
        mem.WriteByte(B)
    Next
    ByteArray = mem.ToArray()
    BinaryStream.Write ByteArray 

    BinaryStream.Position = &HC6 'SET OFFSET
    mem.SetLength(0)
    i=ucase(surname(ln)) 
    ReDim d(&H34) 'd(Len(i)*2)
    For x = 1 To Len(i) 
        d((x-1)*2) = asc(Mid(i,x,1))
    Next
    For Each B in d
        mem.WriteByte(B)
    Next
    ByteArray = mem.ToArray()
    BinaryStream.Write ByteArray 

    BinaryStream.Position = &H13E 'PLAN ID SET OFFSET
    mem.SetLength(0)
    Randomize
    i=(Int((999999999999)*Rnd+9999999999))
    ReDim d(&H13*2) 'd(Len(i)*2)
    For x = 1 To Len(i) 
        d((x-1)*2) = asc(Mid(i,x,1))
    Next
    For Each B in d
        mem.WriteByte(B)
    Next
    ByteArray = mem.ToArray()
    BinaryStream.Write ByteArray 

    BinaryStream.Position = &H1C6 'SET OFFSET
    mem.SetLength(0)
    i=RandomDateYYYYMMDD("2016/03/10", "2024/03/10") & "0050" 'scandate
    ReDim d(Len(i)*2)
    For x = 1 To Len(i) 
        d((x-1)*2) = asc(Mid(i,x,1))
    Next
    For Each B In d
        mem.WriteByte(B)
    Next
    ByteArray = mem.ToArray()
    BinaryStream.Write ByteArray 

    BinaryStream.Position = &H1EC 'DOB SET OFFSET
    mem.SetLength(0)
    i=RandomDateMMDDYYYY("1945/03/10", "2000/03/10") &" " 'birthdate and space
    ReDim d(Len(i))
    For x = 1 To Len(i) 
        d((x-1)) = asc(Mid(i,x,1))
    Next
    For Each B In d
        mem.WriteByte(B)
    Next
    ByteArray = mem.ToArray()
    BinaryStream.Write ByteArray 

    BinaryStream.SaveToFile strFilename, adSaveCreateOverWrite
End Function

Function WriteWildFiles(strFolder, strWild) 
Dim objfs 
Dim objFolder
dim objFiles
Dim objFile
Dim strDesc

Set objfs = CreateObject("Scripting.FileSystemObject")
On error resume next ' Intercept No Folder
    If strFolder ="" then strFolder = objFS.GetAbsolutePathName(".") & "\"
    Set objFolder = objFS.GetFolder(strFolder)
    if Err.Number <> 0 then strDesc = Err.Description
    On error goto 0
    If Len(strDesc) = 0 then
        Set objFiles = objFolder.Files
        For Each objFile in ObjFiles
            if instr(1,objFile.Name, strWild, 1) > 0 then 
                  'WriteWildFiles = strFolder & objFile.Name 
                  OverWritePersonName (strFolder & objFile.Name)
            End if
        Next
    End if 

Set objfs = nothing
End Function
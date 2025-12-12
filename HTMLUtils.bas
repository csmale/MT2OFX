Attribute VB_Name = "HTMLUtils"
Option Explicit

Private xEntInfo As Scripting.Dictionary

Private Sub InitEntities()
    If Not xEntInfo Is Nothing Then
        Exit Sub
    End If
    Set xEntInfo = New Scripting.Dictionary
    xEntInfo("quot") = &H22
    xEntInfo("amp") = &H26
    xEntInfo("apos") = &H27
    xEntInfo("lt") = &H3C
    xEntInfo("gt") = &H3E
    xEntInfo("nbsp") = &HA0
    xEntInfo("iexcl") = &HA1
    xEntInfo("cent") = &HA2
    xEntInfo("pound") = &HA3
    xEntInfo("curren") = &HA4
    xEntInfo("yen") = &HA5
    xEntInfo("brvbar") = &HA6
    xEntInfo("sect") = &HA7
    xEntInfo("uml") = &HA8
    xEntInfo("copy") = &HA9
    xEntInfo("ordf") = &HAA
    xEntInfo("laquo") = &HAB
    xEntInfo("not") = &HAC
    xEntInfo("shy") = &HAD
    xEntInfo("reg") = &HAE
    xEntInfo("macr") = &HAF
    xEntInfo("deg") = &HB0
    xEntInfo("plusmn") = &HB1
    xEntInfo("sup2") = &HB2
    xEntInfo("sup3") = &HB3
    xEntInfo("acute") = &HB4
    xEntInfo("micro") = &HB5
    xEntInfo("para") = &HB6
    xEntInfo("middot") = &HB7
    xEntInfo("cedil") = &HB8
    xEntInfo("sup1") = &HB9
    xEntInfo("ordm") = &HBA
    xEntInfo("raquo") = &HBB
    xEntInfo("frac14") = &HBC
    xEntInfo("frac12") = &HBD
    xEntInfo("frac34") = &HBE
    xEntInfo("iquest") = &HBF
    xEntInfo("Agrave") = &HC0
    xEntInfo("Aacute") = &HC1
    xEntInfo("Acirc") = &HC2
    xEntInfo("Atilde") = &HC3
    xEntInfo("Auml") = &HC4
    xEntInfo("Aring") = &HC5
    xEntInfo("AElig") = &HC6
    xEntInfo("Ccedil") = &HC7
    xEntInfo("Egrave") = &HC8
    xEntInfo("Eacute") = &HC9
    xEntInfo("Ecirc") = &HCA
    xEntInfo("Euml") = &HCB
    xEntInfo("Igrave") = &HCC
    xEntInfo("Iacute") = &HCD
    xEntInfo("Icirc") = &HCE
    xEntInfo("Iuml") = &HCF
    xEntInfo("ETH") = &HD0
    xEntInfo("Ntilde") = &HD1
    xEntInfo("Ograve") = &HD2
    xEntInfo("Oacute") = &HD3
    xEntInfo("Ocirc") = &HD4
    xEntInfo("Otilde") = &HD5
    xEntInfo("Ouml") = &HD6
    xEntInfo("times") = &HD7
    xEntInfo("Oslash") = &HD8
    xEntInfo("Ugrave") = &HD9
    xEntInfo("Uacute") = &HDA
    xEntInfo("Ucirc") = &HDB
    xEntInfo("Uuml") = &HDC
    xEntInfo("Yacute") = &HDD
    xEntInfo("THORN") = &HDE
    xEntInfo("szlig") = &HDF
    xEntInfo("agrave") = &HE0
    xEntInfo("aacute") = &HE1
    xEntInfo("acirc") = &HE2
    xEntInfo("atilde") = &HE3
    xEntInfo("auml") = &HE4
    xEntInfo("aring") = &HE5
    xEntInfo("aelig") = &HE6
    xEntInfo("ccedil") = &HE7
    xEntInfo("egrave") = &HE8
    xEntInfo("eacute") = &HE9
    xEntInfo("ecirc") = &HEA
    xEntInfo("euml") = &HEB
    xEntInfo("igrave") = &HEC
    xEntInfo("iacute") = &HED
    xEntInfo("icirc") = &HEE
    xEntInfo("iuml") = &HEF
    xEntInfo("eth") = &HF0
    xEntInfo("ntilde") = &HF1
    xEntInfo("ograve") = &HF2
    xEntInfo("oacute") = &HF3
    xEntInfo("ocirc") = &HF4
    xEntInfo("otilde") = &HF5
    xEntInfo("ouml") = &HF6
    xEntInfo("divide") = &HF7
    xEntInfo("oslash") = &HF8
    xEntInfo("ugrave") = &HF9
    xEntInfo("uacute") = &HFA
    xEntInfo("ucirc") = &HFB
    xEntInfo("uuml") = &HFC
    xEntInfo("yacute") = &HFD
    xEntInfo("thorn") = &HFE
    xEntInfo("yuml") = &HFF
    xEntInfo("OElig") = &H152
    xEntInfo("oelig") = &H153
    xEntInfo("Scaron") = &H160
    xEntInfo("scaron") = &H161
    xEntInfo("Yuml") = &H178
    xEntInfo("fnof") = &H192
    xEntInfo("circ") = &H2C6
    xEntInfo("tilde") = &H2DC
    xEntInfo("Alpha") = &H391
    xEntInfo("Beta") = &H392
    xEntInfo("Gamma") = &H393
    xEntInfo("Delta") = &H394
    xEntInfo("Epsilon") = &H395
    xEntInfo("Zeta") = &H396
    xEntInfo("Eta") = &H397
    xEntInfo("Theta") = &H398
    xEntInfo("Iota") = &H399
    xEntInfo("Kappa") = &H39A
    xEntInfo("Lambda") = &H39B
    xEntInfo("Mu") = &H39C
    xEntInfo("Nu") = &H39D
    xEntInfo("Xi") = &H39E
    xEntInfo("Omicron") = &H39F
    xEntInfo("Pi") = &H3A0
    xEntInfo("Rho") = &H3A1
    xEntInfo("Sigma") = &H3A3
    xEntInfo("Tau") = &H3A4
    xEntInfo("Upsilon") = &H3A5
    xEntInfo("Phi") = &H3A6
    xEntInfo("Chi") = &H3A7
    xEntInfo("Psi") = &H3A8
    xEntInfo("Omega") = &H3A9
    xEntInfo("alpha") = &H3B1
    xEntInfo("beta") = &H3B2
    xEntInfo("gamma") = &H3B3
    xEntInfo("delta") = &H3B4
    xEntInfo("epsilon") = &H3B5
    xEntInfo("zeta") = &H3B6
    xEntInfo("eta") = &H3B7
    xEntInfo("theta") = &H3B8
    xEntInfo("iota") = &H3B9
    xEntInfo("kappa") = &H3BA
    xEntInfo("lambda") = &H3BB
    xEntInfo("mu") = &H3BC
    xEntInfo("nu") = &H3BD
    xEntInfo("xi") = &H3BE
    xEntInfo("omicron") = &H3BF
    xEntInfo("pi") = &H3C0
    xEntInfo("rho") = &H3C1
    xEntInfo("sigmaf") = &H3C2
    xEntInfo("sigma") = &H3C3
    xEntInfo("tau") = &H3C4
    xEntInfo("upsilon") = &H3C5
    xEntInfo("phi") = &H3C6
    xEntInfo("chi") = &H3C7
    xEntInfo("psi") = &H3C8
    xEntInfo("omega") = &H3C9
    xEntInfo("thetasym") = &H3D1
    xEntInfo("upish") = &H3D2
    xEntInfo("piv") = &H3D6
    xEntInfo("ensp") = &H2002
    xEntInfo("emsp") = &H2003
    xEntInfo("thinsp") = &H2009
    xEntInfo("zwnj") = &H200C
    xEntInfo("zwj") = &H200D
    xEntInfo("lrm") = &H200E
    xEntInfo("rlm") = &H200F
    xEntInfo("ndash") = &H2013
    xEntInfo("mdash") = &H2014
    xEntInfo("lsquo") = &H2018
    xEntInfo("rsquo") = &H2019
    xEntInfo("sbquo") = &H201A
    xEntInfo("ldquo") = &H201C
    xEntInfo("rdquo") = &H201D
    xEntInfo("bdquo") = &H201E
    xEntInfo("dagger") = &H2020
    xEntInfo("Dagger") = &H2021
    xEntInfo("bull") = &H2022
    xEntInfo("hellip") = &H2026
    xEntInfo("permil") = &H2030
    xEntInfo("prime") = &H2032
    xEntInfo("Prime") = &H2033
    xEntInfo("lsaquo") = &H2039
    xEntInfo("rsaquo") = &H203A
    xEntInfo("oline") = &H203E
    xEntInfo("frasl") = &H2044
    xEntInfo("euro") = &H20AC
    xEntInfo("image") = &H2111
    xEntInfo("weierp") = &H2118
    xEntInfo("real") = &H211C
    xEntInfo("trade") = &H2122
    xEntInfo("alefsym") = &H2135
    xEntInfo("larr") = &H2190
    xEntInfo("uarr") = &H2191
    xEntInfo("rarr") = &H2192
    xEntInfo("darr") = &H2193
    xEntInfo("harr") = &H2194
    xEntInfo("crarr") = &H21B5
    xEntInfo("lArr") = &H21D0
    xEntInfo("uArr") = &H21D1
    xEntInfo("rArr") = &H21D2
    xEntInfo("dArr") = &H21D3
    xEntInfo("hArr") = &H21D4
    xEntInfo("forall") = &H2200
    xEntInfo("part") = &H2202
    xEntInfo("exist") = &H2203
    xEntInfo("empty") = &H2205
    xEntInfo("nabla") = &H2207
    xEntInfo("isin") = &H2208
    xEntInfo("notin") = &H2209
    xEntInfo("ni") = &H220B
    xEntInfo("prod") = &H220F
    xEntInfo("sum") = &H2211
    xEntInfo("minus") = &H2212
    xEntInfo("lowast") = &H2217
    xEntInfo("radic") = &H221A
    xEntInfo("prop") = &H221D
    xEntInfo("infin") = &H221E
    xEntInfo("ang") = &H2220
    xEntInfo("and") = &H2227
    xEntInfo("or") = &H2228
    xEntInfo("cap") = &H2229
    xEntInfo("cup") = &H222A
    xEntInfo("int") = &H222B
    xEntInfo("there4") = &H2234
    xEntInfo("sim") = &H223C
    xEntInfo("cong") = &H2245
    xEntInfo("asymp") = &H2248
    xEntInfo("ne") = &H2260
    xEntInfo("equiv") = &H2261
    xEntInfo("le") = &H2264
    xEntInfo("ge") = &H2265
    xEntInfo("sub") = &H2282
    xEntInfo("sup") = &H2283
    xEntInfo("nsub") = &H2284
    xEntInfo("sube") = &H2286
    xEntInfo("supe") = &H2287
    xEntInfo("oplus") = &H2295
    xEntInfo("otimes") = &H2297
    xEntInfo("perp") = &H22A5
    xEntInfo("sdot") = &H22C5
    xEntInfo("lceil") = &H2308
    xEntInfo("rceil") = &H2309
    xEntInfo("lfloor") = &H230A
    xEntInfo("rfloor") = &H230B
    xEntInfo("lang") = &H2329
    xEntInfo("rang") = &H232A
    xEntInfo("loz") = &H25CA
    xEntInfo("spades") = &H2660
    xEntInfo("clubs") = &H2663
    xEntInfo("hearts") = &H2665
    xEntInfo("diams") = &H2666
End Sub

Public Function HTMLDecodeString(sIn As String) As String
    Dim sTmp As String
    Dim iAmp As Long
    Dim iSemi As Long
    Dim sEnt As String
    Dim sVal As String
    
    InitEntities
    
    On Error GoTo baleout
    sTmp = ""
    iSemi = 0
    iAmp = 0
    iAmp = InStr(sIn, "&")
    Do While iAmp > 0
        sTmp = sTmp & Mid$(sIn, iSemi + 1, iAmp - iSemi - 1)
        iSemi = InStr(iAmp + 1, sIn, ";")
        If iSemi = 0 Then ' unencoded & !!!!!!!
            sEnt = "amp"
            iSemi = iAmp
        Else
            sEnt = Mid$(sIn, iAmp + 1, iSemi - iAmp - 1)
            If InStr(sEnt, " ") > 0 Then
                sEnt = "amp"   ' more crap in html, assume unencoded &
                iSemi = iAmp
            End If
        End If
        If Left$(sEnt, 1) = "#" Then
            If Mid$(sEnt, 2, 1) = "x" Then
                sVal = ChrW(CLng("&h" & Mid$(sEnt, 3)))
            ElseIf IsNumeric(Mid$(sEnt, 2)) Then
                sVal = ChrW(CLng(Mid$(sEnt, 2)))
            Else
                Debug.Print "Unable to decode HTML entity " & sEnt
                sVal = "?"
            End If
        Else
            If xEntInfo.Exists(sEnt) Then
                sVal = ChrW(xEntInfo(sEnt))
            Else
                Debug.Print "Unable to decode HTML entity " & sEnt
                sVal = "&" & sEnt & ";" ' leave untranslated - hope someone else can sort it
            End If
        End If
        sTmp = sTmp & sVal
        iAmp = InStr(iSemi + 1, sIn, "&")
    Loop
    sTmp = sTmp & Mid$(sIn, iSemi + 1)
    HTMLDecodeString = sTmp
    Exit Function
baleout:
    HTMLDecodeString = Err.Description
End Function


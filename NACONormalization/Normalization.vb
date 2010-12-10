Imports System.Text.RegularExpressions
'''<summary>
''' Copyright (c) 2000-2002 OCLC Online Computer Library Center, Inc. and other contributors. All rights reserved.
''' The contents of this file, as updated from time to time by the OCLC Office of Research, are subject to OCLC Research Public License Version 2.0 (the "License"); you may not use this file except in compliance with the License. You may obtain a current copy of the License at http://purl.oclc.org/oclc/research/ORPL/.  Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for the specific language governing rights and limitations under the License.  This software consists of voluntary contributions made by many individuals on behalf of OCLC Research. For more information on OCLC Research, please see http://www.oclc.org/oclc/research/.
''' 
''' The Original Code is NACO.java.
''' The Initial Developer of the Original Code is Jenny Toves.
'''
''' The Current Code is Normalization.vb.
''' Current developer: Anthony Gray (yho15524429@yahoo.com)
'''
''' NACO.Data.Normalization methods perform normalizations on input strings according to the Authority File Comparison Rules set forth at http://www.loc.gov/catdir/pcc/naco/normrule.html
'''</summary>
Public Class Normalization

    '* The musical flat is the only character whose Unicode value
    '* differs from its ASCII code. The caller can select which
    '* value they prefer to receive as the musical flat.

    Public Const UNICODE_MUSICAL_FLAT = ChrW(&H266D)
    Public Const MARC_MUSICAL_FLAT = ChrW(&HA9)
    Public Const DEFAULT_SUBFIELD_SEPARATOR = ChrW(&H1F)

    '''<summary>Convert all of the characters in a string according to the Authority File Comparison rules. Uses the Unicode character for the musical flat and uses 0x1F for the subfield delimiter.
    ''' <paramref name="ptr">a UTF-8 string containing subfield delimiters (0x1F), subfield codes and subfield contents. Multiple subfields are expected but the marc tag and indicators are not expected.</paramref>
    ''' <paramref name="leaveFirstComma">flag to leave the first non-trailing comma in the first subfield 'a'.</paramref>
    ''' <paramref name="return">a normalized string</paramref>
    '''</summary>
    Private Function normalize(ByVal ptr As String, ByVal leaveFirstComma As Boolean) As String
        Return NormalizeStandard(ptr, leaveFirstComma, UNICODE_MUSICAL_FLAT, DEFAULT_SUBFIELD_SEPARATOR)
    End Function

    '''<summary>Convert all of the characters in a string according to the Authority File Comparison rules. Uses the Unicode character for the musical flat and uses blank for the subfield delimiter. All commas are converted to blanks.
    ''' <paramref name="ptr">a UTF-8 string containing subfield delimiters (0x1F), subfield codes and subfield contents. Multiple subfields are expected but the marc tag and indicators are not expected.</paramref>
    ''' <paramref name="return">a normalized string</paramref>
    '''</summary>
    Public Shared Function NormalizeSimplified(ByVal ptr As String) As String
        Return NormalizeStandard(ptr, False, "F", " ").Trim
    End Function

    '''<summary>Convert all of the characters in a string according to the Authority File Comparison rules.
    ''' <paramref name="ptr">a UTF-8 string containing subfield delimiters (0x1F), subfield codes and subfield contents. Multiple subfields are expected but the marc tag and indicators are not expected.</paramref>
    ''' <paramref name="leaveFirstComma">flag to leave the first non-trailing comma in THe first subfield 'a'.</paramref>
    ''' <paramref name="musicalFlat">character to use for outputting musical flats</paramref>
    ''' <paramref name="return">a normalized string</paramref>
    ''' '''</summary>
    Public Shared Function NormalizeStandard(ByVal ptr As String, ByVal leaveFirstComma As Boolean, ByVal musicalFlat As Char, ByVal output_Subfield_Separator As Char) As String
        Dim c As Char
        Dim last As Char = " "

        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder()

        Dim commaFound As Boolean = False
        Dim subfieldA As Boolean = True
        Dim firstSFaProcessed As Boolean = False

        Dim ptrChars As Char() = ptr.ToCharArray()
        Dim i As Integer
        Dim firstSF As Integer = 0

        For i = firstSF To ptrChars.Length - 1
            Try
                c = ptrChars(i)

                Select Case c    ' Characters to convert.
                    Case ChrW(&H2070)       ' superscript 0
                        sb.Append("0")
                    Case ChrW(&H2080)       ' subscript 0
                        sb.Append("0")
                    Case ChrW(&HB9)         ' superscript 1
                        sb.Append("1")
                    Case ChrW(&H2071)       ' superscript 1
                        sb.Append("1")
                    Case ChrW(&H2081)       ' subscript 1
                        sb.Append("1")
                    Case ChrW(&HB2)         ' superscript 2
                        sb.Append("2")
                    Case ChrW(&H2072)       ' superscript 2
                        sb.Append("2")
                    Case ChrW(&H2082)       ' subscript 2
                        sb.Append("2")
                    Case ChrW(&HB3)         ' superscript 3
                        sb.Append("3")
                    Case ChrW(&H2073)       ' superscript 3
                        sb.Append("3")
                    Case ChrW(&H2083)       ' subscript 3
                        sb.Append("3")
                    Case ChrW(&H2074)       ' superscript 4
                        sb.Append("4")
                    Case ChrW(&H2084)       ' subscript 4
                        sb.Append("4")
                    Case ChrW(&H2075)       ' superscript 5
                        sb.Append("5")
                    Case ChrW(&H2085)       ' subscript 5
                        sb.Append("5")
                    Case ChrW(&H2076)       ' superscript 6
                        sb.Append("6")
                    Case ChrW(&H2086)       ' subscript 6
                        sb.Append("6")
                    Case ChrW(&H2077)       ' superscript 7
                        sb.Append("7")
                    Case ChrW(&H2087)       ' subscript 7
                        sb.Append("7")
                    Case ChrW(&H2078)       ' superscript 8
                        sb.Append("8")
                    Case ChrW(&H2088)       ' subscript 8
                        sb.Append("8")
                    Case ChrW(&H2079)       ' superscript 9
                        sb.Append("9")
                    Case ChrW(&H2089)       ' subscript 9
                        sb.Append("9")
                    Case ChrW(&H3B1)        ' lc alpha
                        sb.Append("a")
                    Case ChrW(&H391)        ' uc alpha
                        sb.Append("A")
                    Case ChrW(&H3B2)        ' lc beta
                        sb.Append("b")
                    Case ChrW(&H392)        ' uc beta
                        sb.Append("B")
                    Case ChrW(&H3B3)        ' lc gamma
                        sb.Append("g")
                    Case ChrW(&H393)        ' uc gamma
                        sb.Append("G")
                    Case ChrW(&HC6)         ' uc ae digraph
                        sb.Append("ae")
                    Case ChrW(&HE6)         ' lc ae digraph
                        sb.Append("ae")
                    Case ChrW(&H152)        ' uc oe digraph
                        sb.Append("oe")
                    Case ChrW(&H153)        ' lc oe digraph
                        sb.Append("oe")
                    Case ChrW(&H110)        ' uc d with crossbar
                        sb.Append("d")
                    Case ChrW(&H111)        ' lc d with crossbar
                        sb.Append("d")
                    Case ChrW(&HF0)         ' lc eth
                        sb.Append("d")
                    Case ChrW(&H131)        ' lc turkish i
                        sb.Append("i")
                    Case ChrW(&H2113)       ' script small l
                        sb.Append("l")
                    Case ChrW(&H141)        ' uc polish l
                        sb.Append("l")
                    Case ChrW(&H142)        ' lc polish l
                        sb.Append("l")
                    Case ChrW(&H1A0)        ' uc hooked o
                        sb.Append("o")
                    Case ChrW(&H1A1)        ' lc hooked o
                        sb.Append("o")
                    Case ChrW(&HD8)         ' uc scandanavian o
                        sb.Append("o")
                    Case ChrW(&HF8)         ' lc scandanavian o
                        sb.Append("o")
                    Case ChrW(&HDE)         ' uc icelandic thorn
                        sb.Append("th")
                    Case ChrW(&HFE)         ' lc icelanic thorn
                        sb.Append("th")
                    Case ChrW(&H1AF)        ' uc hooked u
                        sb.Append("u")
                    Case ChrW(&H1B0)        ' lc hooked u
                        sb.Append("u")
                    Case ChrW(&H266F)       ' musical sharp
                        sb.Append("#")
                    Case ChrW(&H266D)       ' musical flat
                        sb.Append(musicalFlat)
                    Case "'", "|", "]", "[" ' Characters to delete.
                        sb.Append("")
                    Case ","
                        If (subfieldA And leaveFirstComma And Not commaFound) Then
                            sb.Append(",")
                            commaFound = True
                        Else
                            sb.Append(" ")
                        End If

                    Case "!", "¡", """", "$", "%", "-", ".", "/", "\", ":", ";", "<", "=", ">", "^", "_", "`", "~", "(", ")", "{", "}", "?", "¿", "*", "@", ChrW(&HAE), ChrW(&H2117), ChrW(&HA9), ChrW(&HB1), ChrW(&HA3), ChrW(&H207A), ChrW(&H208A), ChrW(&H207B), ChrW(&H208B), ChrW(&H207D), ChrW(&H207E), ChrW(&H208D), ChrW(&H208E), ChrW(&HBF), ChrW(&HA1), ChrW(&HB0), " "
                        ' *** else drop through and convert to blank
                        If (last <> " " And last <> "\") Then sb.Append(" ")

                    Case ChrW(&H1E), DEFAULT_SUBFIELD_SEPARATOR   ' handle subfield codes
                        ' grab the subfield code
                        Dim n As Integer = i + 1
                        c = ptrChars(n)
                        i += 1

                        ' trim trailing  blanks & commas
                        sb = New Text.StringBuilder(sb.ToString.TrimEnd(" ", ","))

                        If i <> (firstSF + 1) Then
                            If sb.Length > 0 And last = " " Then
                                sb.Append(output_Subfield_Separator)
                            Else
                                sb.Append(output_Subfield_Separator)
                            End If
                        End If

                        ' flag the first subfieldA so that we can handle the comma flag correctly
                        If (Not firstSFaProcessed And c = "a") Then
                            subfieldA = True
                        Else
                            subfieldA = False
                        End If

                    Case Else
                        ' Characters to keep.
                        ' Keep characters in the ASCII range.
                        ' Convert to lowercase.
                        If (c >= ChrW(&H20) And c < ChrW(&H7F)) Then
                            sb.Append(c.ToString.ToLower)
                        End If
                End Select
                If (subfieldA) Then firstSFaProcessed = True
                If (sb.Length() > 0) Then last = sb.ToString.Last

            Catch ex As Exception

            End Try
        Next

        ' Trim off trailing commas and blanks
        Return Regex.Replace(sb.ToString.TrimEnd(" ", ","), "\s+", " ")
    End Function
End Class



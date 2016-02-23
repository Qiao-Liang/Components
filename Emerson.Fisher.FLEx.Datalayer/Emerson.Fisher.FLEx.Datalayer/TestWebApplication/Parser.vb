Imports System
Imports System.IO
Imports System.Collections
Imports System.Collections.Generic

Namespace TestWebApplication

    Public Class Parser

        Private _strTemplateBlock As String
        Private _hstValues As Hashtable
        Private _ErrorMessage As New Hashtable()
        Private _ParsedBlock As String

        Private _Blocks = New Dictionary(Of String, Parser)

        Private VariableTagBegin As String = "##"
        Private VariableTagEnd As String = "##"

        Private ModificatorTag As String = ":"
        Private ModificatorParamSep As String = ","

        Private ConditionTagIfBegin As String = "##If--"
        Private ConditionTagIfEnd As String = "##"
        Private ConditionTagElseBegin As String = "##Else--"
        Private ConditionTagElseEnd As String = "##"
        Private ConditionTagEndIfBegin As String = "##EndIf--"
        Private ConditionTagEndIfEnd As String = "##"

        Private BlockTagBeginBegin As String = "##BlockBegin--"
        Private BlockTagBeginEnd As String = "##"
        Private BlockTagEndBegin As String = "##BlockEnd--"
        Private BlockTagEndEnd As String = "##"

        Public Property TemplateBlock() As String
            Get
                Return Me._strTemplateBlock
            End Get
            Set(ByVal value As String)
                Me._strTemplateBlock = value
                ParseBlocks()
            End Set
        End Property

        Public Property Variables() As Hashtable
            Get
                Return Me._hstValues
            End Get
            Set(ByVal value As Hashtable)
                Me._hstValues = value
            End Set
        End Property

        Public ReadOnly Property ErrorMessage() As Hashtable
            Get
                Return _ErrorMessage
            End Get
        End Property

        ReadOnly Property Blocks() As Dictionary(Of String, Parser)
            Get
                Return _Blocks
            End Get
        End Property

        ' Creates a new instance of TemplateParser
        Public Sub New()
            Me._strTemplateBlock = ""
        End Sub

        Public Sub New(ByVal FilePath As String)
            ReadTemplateFromFile(FilePath)
            ParseBlocks()
        End Sub

        Public Sub New(ByVal Variables As Hashtable)
            Me._hstValues = Variables
        End Sub

        Public Sub New(ByVal FilePath As String, ByVal Variables As Hashtable)
            ReadTemplateFromFile(FilePath)
            Me._hstValues = Variables
            ParseBlocks()
        End Sub

        ' Setup template from specified file
        Public Sub SetTemplateFromFile(ByVal FilePath As String)
            ReadTemplateFromFile(FilePath)
        End Sub

        ' Setup template as string block
        Public Sub SetTemplate(ByVal TemplateBlock As String)
            Me.TemplateBlock = TemplateBlock
        End Sub

        ' Parse template after setuping Template and Variables
        Public Function Parse() As String
            ParseConditions()
            ParseVariables()
            Return Me._ParsedBlock
        End Function

        'Parse Template Block
        Public Function ParseBlock(ByVal BlockName As String, ByVal Variables As Hashtable) As String
            If Not Me._Blocks.ContainsKey(BlockName) Then
                Throw New ArgumentException([String].Format("Could not find Block with Name '{0}'", BlockName))
            End If

            Me._Blocks(BlockName).Variables = Variables
            Return Me._Blocks(BlockName).Parse()
        End Function

        ' Parse template and save result into specified file
        Public Function ParseToFile(ByVal FilePath As String, ByVal ReplaceIfExists As Boolean) As Boolean
            If File.Exists(FilePath) And Not ReplaceIfExists Then
                Return False
            Else
                Dim sr As StreamWriter = File.CreateText(FilePath)
                sr.Write(Parse())
                sr.Close()
                Return True
            End If
        End Function

        ' Read template content from specified file
        Private Sub ReadTemplateFromFile(ByVal FilePath As String)
            If Not File.Exists(FilePath) Then
                Throw New ArgumentException("Template file does not exist.")
            End If

            Dim reader As New StreamReader(FilePath)
            Me.TemplateBlock = reader.ReadToEnd()
            reader.Close()
        End Sub

        ' Parse all blocks in template
        Private Sub ParseBlocks()
            'int idxPrevious = 0;
            Dim idxCurrent As Integer = Me._strTemplateBlock.IndexOf(Me.BlockTagBeginBegin, idxCurrent)
            While (idxCurrent <> -1)
                Dim BlockName As String
                Dim idxBlockBeginBegin, idxBlockBeginEnd, idxBlockEndBegin As Integer

                idxBlockBeginBegin = idxCurrent
                idxCurrent += Me.BlockTagBeginBegin.Length

                ' Searching for BlockBeginEnd Index
                idxBlockBeginEnd = Me._strTemplateBlock.IndexOf(Me.BlockTagBeginEnd, idxCurrent)
                If idxBlockBeginEnd = -1 Then
                    Throw New Exception("Could not find BlockTagBeginEnd")
                End If
                ' Getting Block Name
                BlockName = Me._strTemplateBlock.Substring(idxCurrent, idxBlockBeginEnd - idxCurrent)
                idxCurrent = idxBlockBeginEnd + Me.BlockTagBeginEnd.Length

                ' Getting End of Block index
                Dim EndBlockStatment As String = Me.BlockTagEndBegin + BlockName + Me.BlockTagEndEnd
                idxBlockEndBegin = Me._strTemplateBlock.IndexOf(EndBlockStatment, idxCurrent)
                If idxBlockEndBegin = -1 Then
                    Throw New Exception("Could not find End of Block with name '" + BlockName + "'")
                End If
                ' Add Block to Dictionary
                Dim block As New Parser()
                block.TemplateBlock = Me._strTemplateBlock.Substring(idxCurrent, idxBlockEndBegin - idxCurrent)
                Me._Blocks.Add(BlockName, block)

                ' Remove Block Declaration From Template
                Me._strTemplateBlock = Me._strTemplateBlock.Remove(idxBlockBeginBegin, idxBlockEndBegin - idxBlockBeginBegin)

                idxCurrent = idxBlockBeginBegin
                idxCurrent = Me._strTemplateBlock.IndexOf(Me.BlockTagBeginBegin, idxCurrent)
            End While
        End Sub

        ' Parse all conditions in template
        Private Sub ParseConditions()
            Dim idxPrevious As Integer = 0
            Dim idxCurrent As Integer = Me._strTemplateBlock.IndexOf(Me.ConditionTagIfBegin, idxCurrent)
            Me._ParsedBlock = ""
            While (idxCurrent <> -1)
                Dim VarName As String
                Dim TrueBlock, FalseBlock As String
                Dim ElseStatment, EndIfStatment As String
                Dim idxIfBegin, idxIfEnd, idxElseBegin, idxEndIfBegin As Integer
                Dim boolValue As Boolean

                idxIfBegin = idxCurrent
                idxCurrent += Me.ConditionTagIfBegin.Length

                ' Searching for EndIf Index
                idxIfEnd = Me._strTemplateBlock.IndexOf(Me.ConditionTagIfEnd, idxCurrent)
                If idxIfEnd = -1 Then
                    Throw New Exception("Could not find ConditionTagIfEnd")
                End If
                ' Getting Value Name
                VarName = Me._strTemplateBlock.Substring(idxCurrent, idxIfEnd - idxCurrent)

                idxCurrent = idxIfEnd + Me.ConditionTagIfEnd.Length

                ' Compare ElseIf and EndIf Indexes
                ElseStatment = Me.ConditionTagElseBegin + VarName + Me.ConditionTagElseEnd
                EndIfStatment = Me.ConditionTagEndIfBegin + VarName + Me.ConditionTagEndIfEnd
                idxElseBegin = Me._strTemplateBlock.IndexOf(ElseStatment, idxCurrent)
                idxEndIfBegin = Me._strTemplateBlock.IndexOf(EndIfStatment, idxCurrent)
                If idxElseBegin > idxEndIfBegin Then
                    Throw New Exception("Condition Else Tag placed after Condition Tag EndIf for '" + VarName + "'")
                End If
                ' Getting True and False Condition Blocks
                If idxElseBegin <> -1 Then
                    TrueBlock = Me._strTemplateBlock.Substring(idxCurrent, idxElseBegin - idxCurrent)
                    FalseBlock = Me._strTemplateBlock.Substring(idxElseBegin + ElseStatment.Length, idxEndIfBegin - idxElseBegin - ElseStatment.Length)
                Else
                    TrueBlock = Me._strTemplateBlock.Substring(idxCurrent, idxEndIfBegin - idxCurrent)
                    FalseBlock = ""
                End If

                ' Parse Condition
                Try
                    boolValue = Convert.ToBoolean(Me._hstValues(VarName))
                Catch
                End Try

                Dim BeforeBlock As String = Me._strTemplateBlock.Substring(idxPrevious, idxIfBegin - idxPrevious)

                If Me._hstValues.ContainsKey(VarName) And boolValue Then
                    Me._ParsedBlock += BeforeBlock + TrueBlock.Trim()
                Else
                    Me._ParsedBlock += BeforeBlock + FalseBlock.Trim()
                End If

                idxCurrent = idxEndIfBegin + EndIfStatment.Length
                idxPrevious = idxCurrent

                idxCurrent = Me._strTemplateBlock.IndexOf(Me.ConditionTagIfBegin, idxCurrent)
            End While
            Me._ParsedBlock += Me._strTemplateBlock.Substring(idxPrevious)
        End Sub

        ' Parse all variables in template
        Private Sub ParseVariables()
            Dim idxCurrent As Integer = Me._ParsedBlock.IndexOf(Me.VariableTagBegin, idxCurrent)
            While (idxCurrent <> -1)
                Dim VarName, VarValue As String
                Dim idxVarTagEnd As Integer

                idxVarTagEnd = Me._ParsedBlock.IndexOf(Me.VariableTagEnd, idxCurrent + Me.VariableTagBegin.Length)
                If idxVarTagEnd = -1 Then
                    Throw New Exception([String].Format("Index {0}: could not find Variable End Tag", idxCurrent))
                End If
                ' Getting Variable Name
                VarName = Me._ParsedBlock.Substring(idxCurrent + Me.VariableTagBegin.Length, idxVarTagEnd - idxCurrent - Me.VariableTagBegin.Length)

                ' Checking for Modificators
                Dim VarParts As String() = VarName.Split(Me.ModificatorTag.ToCharArray())
                VarName = VarParts(0)

                ' Getting Variable Value
                ' If Variable doesn't exist in _hstValue then
                ' Variable Value equal empty string
                ' [added 6/6/2006] If variable is null than it will also has empty string
                VarValue = [String].Empty
                If Me._hstValues.ContainsKey(VarName) And Not (Me._hstValues(VarName) Is Nothing) Then
                    VarValue = Me._hstValues(VarName).ToString()
                End If

                ' Apply All Modificators to Variable Value
                Dim i As Integer
                For i = 1 To VarParts.Length - 1
                    Me.ApplyModificator(VarValue, VarParts(i))
                Next i
                ' Replace Variable in Template
                Me._ParsedBlock = Me._ParsedBlock.Substring(0, idxCurrent) + VarValue + Me._ParsedBlock.Substring((idxVarTagEnd + Me.VariableTagEnd.Length))

                ' Add Length of added value to Current index 
                ' to prevent looking for variables in the added value
                ' Fixed Date: April 5, 2006
                idxCurrent += VarValue.Length

                idxCurrent = Me._ParsedBlock.IndexOf(Me.VariableTagBegin, idxCurrent)
            End While
        End Sub

        ' Method for applying modificators to variable value
        Private Sub ApplyModificator(ByRef Value As String, ByVal Modificator As String)
            ' Checking for parameters
            Dim strModificatorName As String = ""
            Dim strParameters As String = ""
            Dim idxStartBrackets, idxEndBrackets As Integer
            idxStartBrackets = Modificator.IndexOf("(")
            If idxStartBrackets <> -1 Then
                idxEndBrackets = Modificator.IndexOf(")", idxStartBrackets)
                If idxEndBrackets = -1 Then
                    Throw New Exception("Incorrect modificator expression")
                Else
                    strModificatorName = Modificator.Substring(0, idxStartBrackets).ToUpper()
                    strParameters = Modificator.Substring(idxStartBrackets + 1, idxEndBrackets - idxStartBrackets - 1)
                End If
            Else
                strModificatorName = Modificator.ToUpper()
            End If
            Dim arrParameters As String() = strParameters.Split(Me.ModificatorParamSep.ToCharArray())
            Dim i As Integer
            For i = 0 To arrParameters.Length - 1
                arrParameters(i) = arrParameters(i).Trim()
            Next i
            Try
                Dim typeModificator As Type = Type.GetType(("TemplateParser.Modificators." + strModificatorName))
                If typeModificator.IsSubclassOf(Type.GetType("TemplateParser.Modificators.Modificator")) Then
                    Dim objModificator As Modificators.Modificator = CType(Activator.CreateInstance(typeModificator), Modificators.Modificator)
                    objModificator.Apply(Value, arrParameters)
                End If
            Catch
            End Try
        End Sub

    End Class

End Namespace
Imports System
Imports System.Collections

Namespace TestWebApplication.Modificators

    ' Abstract class for Modificators
    Public MustInherit Class Modificator
        Protected _parameters As New Hashtable()

        Public ReadOnly Property Parameters() As Hashtable
            Get
                Return _parameters
            End Get
        End Property

        Public MustOverride Sub Apply(ByRef Value As String, ByVal ParamArray Parameters() As String)
    End Class

    Class NL2BR
        Inherits Modificator

        Public Overrides Sub Apply(ByRef Value As String, ByVal ParamArray Parameters() As String)
            Value = Value.Replace(ControlChars.Lf, "<br>")
        End Sub
    End Class

    Class HTMLENCODE
        Inherits Modificator

        Public Overrides Sub Apply(ByRef Value As String, ByVal ParamArray Parameters() As String)
            Value = Value.Replace("&", "&amp;")
            Value = Value.Replace("<", "&lt;")
            Value = Value.Replace(">", "&gt;")
        End Sub
    End Class

    Class UPPER
        Inherits Modificator

        Public Overrides Sub Apply(ByRef Value As String, ByVal ParamArray Parameters() As String)
            Value = Value.ToUpper()
        End Sub
    End Class

    Class LOWER
        Inherits Modificator

        Public Overrides Sub Apply(ByRef Value As String, ByVal ParamArray Parameters() As String)
            Value = Value.ToLower()
        End Sub
    End Class

    Class TRIM
        Inherits Modificator

        Public Overrides Sub Apply(ByRef Value As String, ByVal ParamArray Parameters() As String)
            Value = Value.Trim()
        End Sub
    End Class

    Class TRIMEND
        Inherits Modificator

        Public Overrides Sub Apply(ByRef Value As String, ByVal ParamArray Parameters() As String)
            Value = Value.TrimEnd()
        End Sub
    End Class

    Class TRIMSTART
        Inherits Modificator

        Public Overrides Sub Apply(ByRef Value As String, ByVal ParamArray Parameters() As String)
            Value = Value.TrimStart()
        End Sub
    End Class

    Class [DEFAULT]
        Inherits Modificator

        Public Overrides Sub Apply(ByRef Value As String, ByVal ParamArray Parameters() As String)
            If Value Is Nothing Or Value.Trim() = String.Empty Then
                Value = Parameters(0)
            End If
        End Sub
    End Class

End Namespace
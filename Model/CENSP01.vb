
Option Explicit On

Public Class CENSP01

    '
    '   CENSP01.cls
    '
    '   A class that provides a simple container for a Fasteners to be processed
    '
    '   Author: Stacey Keller
    '   Date:   6/16/10
    '

    Private Name_ As String
    Private X_ As Double
    Private Y_ As Double
    Private Z_ As Double
    'comment

    Public Sub New()
        Name_ = ""
        X = 0
        Y = 0
        Z = 0
    End Sub



    Public Property X()
        Get
            X = X_
        End Get
        Set
            X_ = value
        End Set
    End Property
    Public Property Y()
        Get
            Y = Y_
        End Get
        Set
            Y_ = value
        End Set
    End Property
    Public Property Z()
        Get
            Z = Z_
        End Get
        Set
            Z_ = value
        End Set
    End Property

    Public Property name()
        Get
            name = Name_
        End Get
        Set
            Name_ = value
        End Set
    End Property
    Public Function to_point() As CENPoint
        Dim aaa As New CENPoint()
        With aaa
            .X = X
            .Y = Y
            .Z = Z
            .FastenerName = name
        End With
        aaa.iniuuid()
        Return aaa

    End Function

End Class

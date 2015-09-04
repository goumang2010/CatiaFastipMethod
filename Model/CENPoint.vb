Option Explicit On

Imports ProductStructureTypeLib
Imports MECMOD
'Imports PARTITF
Imports HybridShapeTypeLib
Public Class CENPoint

    '   CENPoint.cls
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
    Private MyPoint_ As HybridShape
    Private MyVector_ As HybridShape
    '可以用这个表示processtype

    Private Diam_ As String
    Private Frame_ As String
    Private FastenerName_ As String
    Private PointName_ As String
    Private uuid_ As String
    Public uuidP As String
    Public xxr As Decimal
    Public yyr As Decimal
    Public zzr As Decimal
    Public PFname As String = ""
    'comment

    Public Sub New()
        Name_ = ""
        Frame_ = ""
        uuidP = ""
        Diam_ = ""
        MyPoint_ = Nothing
        MyVector_ = Nothing
        X_ = 0
        Y_ = 0
        Z_ = 0
        FastenerName_ = ""
    End Sub




    Public Sub update(Name As String, Frame As String, Diam As String, ByRef MyPoint As Object, ByRef MyVector As Object, X As Double, Y As Double, Z As Double, FastenerName As String, pointname As String)
        Name_ = Name
        Frame_ = Frame
        Diam_ = Diam
        MyPoint_ = MyPoint
        MyVector_ = MyVector
        X_ = X
        Y_ = Y
        Z_ = Z
        FastenerName_ = FastenerName
        PointName_ = pointname
        iniuuid()
    End Sub

    Public Sub iniuuid()

   


        xxr = Math.Round(X, 2)
        yyr = Math.Round(Y, 2)
        zzr = Math.Round(Z, 2)
        uuid_ = xxr.ToString + "_" + yyr.ToString + "_" + zzr.ToString

        'xxr = Math.Round(X / 10, 0) * 10
        'yyr = Math.Round(Y / 10, 0) * 10
        'zzr = Math.Round(Z / 10, 0) * 10
        'uuid = xxr.ToString + "_" + yyr.ToString + "_" + zzr.ToString
    End Sub

    '对于比较紧固件
    'Public Sub iniuuid2()

    '    Dim xxr As Integer
    '    Dim yyr As Integer
    '    Dim zzr As Integer


    '    xxr = Math.Round(X / 10, 0) * 10
    '    yyr = Math.Round(Y / 10, 0) * 10
    '    zzr = Math.Round(Z / 10, 0) * 10


    '    uuid = xxr.ToString + "_" + yyr.ToString + "_" + zzr.ToString
    'End Sub
    Public Property uuid() As String
        Get
            uuid = uuid_
        End Get
        Set(value As String)

            uuid_ = value
        End Set

    End Property
    Public Property X()
        Get
            X = X_
        End Get
        Set(value)
            X_ = value
        End Set
    End Property
    Public Property Y()
        Get
            Y = Y_
        End Get
        Set(value)
            Y_ = value
        End Set
    End Property
    Public Property Z()
        Get
            Z = Z_
        End Get
        Set(value)
            Z_ = value
        End Set
    End Property

    Public Property name()
        Get
            name = Name_
        End Get
        Set(value)
            Name_ = value
        End Set
    End Property
    Public Property MyPoint()
        Get
            MyPoint = MyPoint_
        End Get
        Set(value)
            MyPoint_ = value
        End Set
    End Property
    Public Function getpoint() As HybridShape
        Return MyPoint_
    End Function
    Public Function getvector() As HybridShape
        Return MyVector_
    End Function

    Public Property MyVector()
        Get
            MyVector = MyVector_
        End Get
        Set(value)
            MyVector_ = value
        End Set
    End Property

    Public Property Diam()
        Get
            Diam = Diam_
        End Get
        Set(value)
            Diam_ = value
        End Set
    End Property
    Public Property PointName()
        Get
            PointName = PointName_
        End Get
        Set(value)
            PointName_ = value
        End Set
    End Property
    Public Property Frame()
        Get
            Frame = Frame_
        End Get
        Set(value)
            Frame_ = value
        End Set
    End Property
    Public Property FastenerName()
        Get
            FastenerName = FastenerName_
        End Get
        Set(value)
            FastenerName_ = value
        End Set
    End Property




End Class

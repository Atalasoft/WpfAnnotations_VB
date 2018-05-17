Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Windows
Imports System.Runtime.Serialization
Imports System.Windows.Media
Imports Atalasoft.Annotate
Imports Atalasoft.Annotate.Wpf.Renderer
Imports Atalasoft.Annotate.Wpf

''' <summary>
''' This data class is used only for serialization in WPF.
''' </summary>
<Serializable()> _
Public Class TriangleData
    Inherits PointBaseData
    Private _fill As AnnotationBrush = WpfAnnotationUI.DefaultFill
    Private _outline As AnnotationPen = WpfAnnotationUI.DefaultOutline

    Public Sub New()
    End Sub

    Public Sub New(ByVal points() As Point, ByVal fill As AnnotationBrush, ByVal outline As AnnotationPen)
        MyBase.New(New PointFCollection(WpfObjectConverter.ConvertPointF(points)))
        _fill = fill
        _outline = outline
    End Sub

    Public Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        MyBase.New(info, context)
        _fill = CType(info.GetValue("Fill", GetType(AnnotationBrush)), AnnotationBrush)
        _outline = CType(info.GetValue("Outline", GetType(AnnotationPen)), AnnotationPen)
    End Sub

    Public Property Fill() As AnnotationBrush
        Get
            Return _fill
        End Get
        Set(ByVal value As AnnotationBrush)
            _fill = value
        End Set
    End Property

    Public Property Outline() As AnnotationPen
        Get
            Return _outline
        End Get
        Set(ByVal value As AnnotationPen)
            _outline = value
        End Set
    End Property

    Public Overrides Sub GetObjectData(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        MyBase.GetObjectData(info, context)
        info.AddValue("Fill", _fill)
        info.AddValue("Outline", _outline)
    End Sub

    Public Overrides Function Clone() As Object
        Dim data As New TriangleData()
        MyBase.CloneBaseData(data)

        If _fill IsNot Nothing Then
            data._fill = _fill.Clone()
        End If
        If _outline IsNot Nothing Then
            data._outline = _outline.Clone()
        End If

        Return data
    End Function
End Class

''' <summary>
''' This class is the actual annotation that will be used in WPF.
''' </summary>
''' <remarks>
''' Don't forget to add a UI factory for this annotation. An example would be:
''' Me.AnnotationViewer.Annotations.Factories.Add(New WpfAnnotationUIFactory(Of TriangleAnnotation, TriangleData)())
''' </remarks>
Public Class TriangleAnnotation
    Inherits WpfPointBaseAnnotation
    Shared Sub New()
        WpfAnnotationRenderers.Add(GetType(TriangleAnnotation), New TriangleAnnotationRenderingEngine())
    End Sub

    Public Sub New()
        Me.New(Nothing, WpfAnnotationUI.DefaultFill, WpfAnnotationUI.DefaultOutline)
    End Sub

    Public Sub New(ByVal points() As Point, ByVal fill As AnnotationBrush, ByVal outline As AnnotationPen)
        MyBase.New(3, points)
        ' WpfAnnotationUI already has dependency properties for fill and outline.
        Me.SetValue(FillProperty, fill)
        Me.SetValue(OutlineProperty, outline)
    End Sub

    Public Sub New(ByVal data As TriangleData)
        MyBase.New(3, data)
        Me.SetValue(FillProperty, data.Fill)
        Me.SetValue(OutlineProperty, data.Outline)
    End Sub

    Public Property Fill() As AnnotationBrush
        Get
            Return CType(GetValue(FillProperty), AnnotationBrush)
        End Get
        Set(ByVal value As AnnotationBrush)
            SetValue(FillProperty, value)
        End Set
    End Property

    Public Property Outline() As AnnotationPen
        Get
            Return CType(GetValue(OutlineProperty), AnnotationPen)
        End Get
        Set(ByVal value As AnnotationPen)
            SetValue(OutlineProperty, value)
        End Set
    End Property

    Public Overrides ReadOnly Property SupportsMultiClickCreation() As Boolean
        Get
            Return True
        End Get
    End Property

    ' This is used for rendering and hit testing.
    Public Overrides ReadOnly Property Geometry() As Geometry
        Get
            Return WpfPointBaseAnnotation.GeometryFromPoints(Me.Points.ToArray(), WpfFreehandLineType.Straight, True, (Me.Fill IsNot Nothing))
        End Get
    End Property

    Protected Overrides Function CloneOverride(ByVal clone As WpfAnnotationUI) As WpfAnnotationUI
        Dim ann As TriangleAnnotation = TryCast(clone, TriangleAnnotation)
        If ann Is Nothing Then
            Return Nothing
        End If

        Dim fill As AnnotationBrush = Me.Fill
        ann.Fill = (If(fill Is Nothing, Nothing, fill.Clone()))

        Dim outline As AnnotationPen = Me.Outline
        ann.Outline = (If(outline Is Nothing, Nothing, outline.Clone()))

        If Me.Points.Count > 0 Then
            ann.Points.AddRange(Me.Points.ToArray())
        End If

        ann.GripMode = Me.GripMode

        Return ann
    End Function

    ' This is used when serializing the annotation to XMP.
    Protected Overrides Function CreateDataSnapshotOverride() As AnnotationData
        Dim points() As Point = Me.Points.ToArray()
        Dim fill As AnnotationBrush = Me.Fill
        Dim outline As AnnotationPen = Me.Outline

        If fill IsNot Nothing Then
            fill = fill.Clone()
        End If
        If outline IsNot Nothing Then
            outline = outline.Clone()
        End If

        Return New TriangleData(points, fill, outline)
    End Function

End Class

Public Class TriangleAnnotationRenderingEngine
    Inherits WpfAnnotationRenderingEngine(Of TriangleAnnotation)
    Public Sub New()
    End Sub

    Protected Overrides Sub OnRenderAnnotation(ByVal annotation As TriangleAnnotation, ByVal environment As WpfRenderEnvironment)
        Dim pen As Pen = WpfObjectConverter.ConvertAnnotationPen(annotation.Outline)
        Dim brush As Brush = WpfObjectConverter.ConvertAnnotationBrush(annotation.Fill)
        environment.DrawingContext.DrawGeometry(brush, pen, annotation.Geometry)
    End Sub
End Class

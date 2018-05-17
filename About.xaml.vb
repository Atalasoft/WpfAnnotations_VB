Partial Public Class About

    ''' <summary>
    ''' This constructor tries to mirror the one used in the WinformsDemos
    ''' </summary>
    ''' <param name="dialogLabel"></param>
    ''' <param name="demoName"></param>
    Public Sub New(ByVal dialogLabel As String, ByVal demoName As String)
        InitializeComponent()

        ' This is the window Title
        Title = dialogLabel

        ' This is the name of the demo at the top of the About Page
        DemoNameField.Content = demoName

        ' This is the Demo Gallery Link Label
        SetDemoGalleryLinkLabel(demoName)
    End Sub

#Region "Alt Constructors"
    ''' <summary>
    ''' Simplified constructor for when you just want mostly defaults
    ''' </summary>
    ''' <param name="demoName"></param>
    Public Sub New(ByVal demoName As String)
        InitializeComponent()

        ' This is the window Title
        Title = "About " + demoName

        ' This is the name of the demo at the top of the About Page
        DemoNameField.Content = demoName

        ' This is the Demo Gallery Link Label
        SetDemoGalleryLinkLabel(demoName)
    End Sub

    ''' <summary>
    ''' An all-in-one version of the constructor for when you want to set everything
    ''' </summary>
    ''' <param name="demoName">Name of the Demo</param>
    ''' <param name="demoLink">Link to the demo's home page</param>
    ''' <param name="demoDesc">Full Description of the demo (separate graphs with two crlf</param>
    Public Sub New(ByVal demoName As String, ByVal demoLink As String, ByVal demoDesc As String)
        Me.New("About " + demoName, demoName, demoLink, demoDesc)
    End Sub

    ''' <summary>
    ''' Full-monte - when you want to set all the params up front in the constructor
    ''' </summary>
    ''' <param name="dialogDesc">The title for the about window</param>
    ''' <param name="demoName">Name of the Demo</param>
    ''' <param name="demoLink">Link to the demo's home page</param>
    ''' <param name="demoDesc">Full Description of the demo (separate graphs with two crlf</param>
    Public Sub New(ByVal dialogDesc As String, ByVal demoName As String, ByVal demoLink As String, ByVal demoDesc As String)
        InitializeComponent()

        ' This is the window Title
        Title = dialogDesc

        ' This is the name of the demo at the top of the About Page
        DemoNameField.Content = demoName

        ' This is the Demo Gallery Link Label
        SetDemoGalleryLinkLabel(demoName)

        ' set the hyperlink
        SetDemoGalleryLink(demoLink)

        ' set the description
        SetDescription(demoDesc)
    End Sub

    Public Sub New()
        InitializeComponent()
    End Sub
#End Region


#Region "Properties"
    Public Property Link() As String
        Get
            Return Me.demoGalleryLink.NavigateUri.ToString()
        End Get
        Set(ByVal value As String)
            SetDemoGalleryLink(value)
        End Set
    End Property

    Public Property Description() As String
        Get
            Return Me.DemoDescription.Text
        End Get
        Set(ByVal value As String)
            SetDescription(value)
        End Set
    End Property
#End Region


#Region "Private Methods"
    Private Sub SetDemoGalleryLinkLabel(ByVal demoName As String)
        ' This is the Demo Gallery Link Label
        If demoName.Contains("Demo") Then
            demoGalleryLinkLabel.Text = demoName + " Home"
        Else
            demoGalleryLinkLabel.Text = demoName + " Demo Home"
        End If
    End Sub
#End Region


#Region "Public Methods"
    Public Sub SetDemoGalleryLink(ByVal uri As String)
        If Not uri.Contains("http://") AndAlso Not uri.Contains("https://") Then
            demoGalleryLink.NavigateUri = New Uri("http://" + uri)
        Else
            demoGalleryLink.NavigateUri = New Uri(uri)
        End If
    End Sub

    Public Sub SetDescription(ByVal desc As String)
        Me.DemoDescription.Text = desc
    End Sub
#End Region


#Region "Event Handlers"
    Private Sub Hyperlink_RequestNavigate(ByVal sender As System.Object, ByVal e As System.Windows.Navigation.RequestNavigateEventArgs)
        If e.Uri.OriginalString <> "" Then
            Process.Start(e.Uri.OriginalString)
        End If
    End Sub
#End Region
End Class

Partial Class Menu
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Requerido para la compatibilidad con el Diseñador de composiciones de clases Windows.Forms
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'El Diseñador de componentes requiere esta llamada.
        InitializeComponent()

    End Sub

    'Component reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de componentes
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de componentes requiere el siguiente procedimiento
    'Se puede modificar usando el Diseñador de componentes.
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.menuWorkSpace = Me.Factory.CreateRibbonTab
        Me.actionBox = Me.Factory.CreateRibbonGroup
        Me.bnHelloWorld = Me.Factory.CreateRibbonButton
        Me.menuWorkSpace.SuspendLayout()
        Me.actionBox.SuspendLayout()
        Me.SuspendLayout()
        '
        'menuWorkSpace
        '
        Me.menuWorkSpace.Groups.Add(Me.actionBox)
        Me.menuWorkSpace.Label = "Hola Mundo"
        Me.menuWorkSpace.Name = "menuWorkSpace"
        '
        'actionBox
        '
        Me.actionBox.Items.Add(Me.bnHelloWorld)
        Me.actionBox.Label = "Actions"
        Me.actionBox.Name = "actionBox"
        '
        'bnHelloWorld
        '
        Me.bnHelloWorld.Label = "Hello World!"
        Me.bnHelloWorld.Name = "bnHelloWorld"
        '
        'Menu
        '
        Me.Name = "Menu"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.menuWorkSpace)
        Me.menuWorkSpace.ResumeLayout(False)
        Me.menuWorkSpace.PerformLayout()
        Me.actionBox.ResumeLayout(False)
        Me.actionBox.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents menuWorkSpace As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents actionBox As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents bnHelloWorld As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Menu() As Menu
        Get
            Return Me.GetRibbon(Of Menu)()
        End Get
    End Property
End Class

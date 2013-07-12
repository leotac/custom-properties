Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports DocumentFormat.OpenXml.CustomProperties
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.VariantTypes

Public Class FilePropDemo
    Inherits System.Windows.Forms.Form

    Private m_fOpenedReadOnly As Boolean
    Private filenames As List(Of String)

    Const c_strFilter As String = "Office Document Files|*.docx;*.docm;*.dotx;*.xlsx;*.xlsm;*.xla;*.xlam;*.pptx;*.pptm;*.vsd|All Files (*.*)|*.*"
    Public Sub New()
        MyBase.New()
        InitializeComponent()
    End Sub
#Region "Windows Form Designer Code"
    Private components As System.ComponentModel.IContainer
    ' Main Form controls...
    Friend WithEvents cmdOpen As System.Windows.Forms.Button
    ' Summary Page Controls...
    ' Statistics Page Controls...
    ' Custom Page Controls...



    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Don't edit this function, the desginer will overwrite this automatically.
    ' Custom settings should be applied in separate function...
    Friend WithEvents lbFileName As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdOpen = New System.Windows.Forms.Button()
        Me.lbFileName = New System.Windows.Forms.Label()
        Me.PropTabs = New System.Windows.Forms.TabControl()
        Me.CustomTab = New System.Windows.Forms.TabPage()
        Me.cboxCustType = New System.Windows.Forms.ComboBox()
        Me.lbCustType = New System.Windows.Forms.Label()
        Me.lbCustValue = New System.Windows.Forms.Label()
        Me.lbCustName = New System.Windows.Forms.Label()
        Me.lbCustNote = New System.Windows.Forms.Label()
        Me.txtCustValue = New System.Windows.Forms.TextBox()
        Me.txtCustName = New System.Windows.Forms.TextBox()
        Me.cmdRemove = New System.Windows.Forms.Button()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.CustListView = New System.Windows.Forms.ListView()
        Me.CustNameCol = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.CustValueCol = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.CustTypeCol = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.CustCountCol = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.lbCustIntro = New System.Windows.Forms.Label()
        Me.PropTabs.SuspendLayout()
        Me.CustomTab.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdOpen
        '
        Me.cmdOpen.Location = New System.Drawing.Point(368, 16)
        Me.cmdOpen.Name = "cmdOpen"
        Me.cmdOpen.Size = New System.Drawing.Size(80, 24)
        Me.cmdOpen.TabIndex = 0
        Me.cmdOpen.Text = "Open..."
        '
        'lbFileName
        '
        Me.lbFileName.Location = New System.Drawing.Point(12, 16)
        Me.lbFileName.Name = "lbFileName"
        Me.lbFileName.Size = New System.Drawing.Size(348, 24)
        Me.lbFileName.TabIndex = 2
        Me.lbFileName.Text = "[Click Open button to read properties from file...]"
        '
        'PropTabs
        '
        Me.PropTabs.Controls.Add(Me.CustomTab)
        Me.PropTabs.Enabled = False
        Me.PropTabs.ItemSize = New System.Drawing.Size(120, 18)
        Me.PropTabs.Location = New System.Drawing.Point(8, 46)
        Me.PropTabs.Name = "PropTabs"
        Me.PropTabs.SelectedIndex = 0
        Me.PropTabs.Size = New System.Drawing.Size(459, 362)
        Me.PropTabs.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
        Me.PropTabs.TabIndex = 1
        '
        'CustomTab
        '
        Me.CustomTab.Controls.Add(Me.cboxCustType)
        Me.CustomTab.Controls.Add(Me.lbCustType)
        Me.CustomTab.Controls.Add(Me.lbCustValue)
        Me.CustomTab.Controls.Add(Me.lbCustName)
        Me.CustomTab.Controls.Add(Me.lbCustNote)
        Me.CustomTab.Controls.Add(Me.txtCustValue)
        Me.CustomTab.Controls.Add(Me.txtCustName)
        Me.CustomTab.Controls.Add(Me.cmdRemove)
        Me.CustomTab.Controls.Add(Me.cmdAdd)
        Me.CustomTab.Controls.Add(Me.CustListView)
        Me.CustomTab.Controls.Add(Me.lbCustIntro)
        Me.CustomTab.Location = New System.Drawing.Point(4, 22)
        Me.CustomTab.Name = "CustomTab"
        Me.CustomTab.Size = New System.Drawing.Size(451, 336)
        Me.CustomTab.TabIndex = 1
        Me.CustomTab.Text = "Custom"
        '
        'cboxCustType
        '
        Me.cboxCustType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboxCustType.Items.AddRange(New Object() {"String", "Integer", "Double", "Boolean", "Date"})
        Me.cboxCustType.Location = New System.Drawing.Point(264, 248)
        Me.cboxCustType.Name = "cboxCustType"
        Me.cboxCustType.Size = New System.Drawing.Size(144, 21)
        Me.cboxCustType.TabIndex = 3
        '
        'lbCustType
        '
        Me.lbCustType.Location = New System.Drawing.Point(224, 248)
        Me.lbCustType.Name = "lbCustType"
        Me.lbCustType.Size = New System.Drawing.Size(40, 16)
        Me.lbCustType.TabIndex = 4
        Me.lbCustType.Text = "Type:"
        '
        'lbCustValue
        '
        Me.lbCustValue.Location = New System.Drawing.Point(16, 280)
        Me.lbCustValue.Name = "lbCustValue"
        Me.lbCustValue.Size = New System.Drawing.Size(40, 16)
        Me.lbCustValue.TabIndex = 5
        Me.lbCustValue.Text = "Value:"
        '
        'lbCustName
        '
        Me.lbCustName.Location = New System.Drawing.Point(16, 248)
        Me.lbCustName.Name = "lbCustName"
        Me.lbCustName.Size = New System.Drawing.Size(40, 16)
        Me.lbCustName.TabIndex = 6
        Me.lbCustName.Text = "Name:"
        '
        'lbCustNote
        '
        Me.lbCustNote.Location = New System.Drawing.Point(16, 208)
        Me.lbCustNote.Name = "lbCustNote"
        Me.lbCustNote.Size = New System.Drawing.Size(216, 32)
        Me.lbCustNote.TabIndex = 7
        Me.lbCustNote.Text = "To add a new item, fill in the information below and click the Add button."
        '
        'txtCustValue
        '
        Me.txtCustValue.Location = New System.Drawing.Point(64, 280)
        Me.txtCustValue.Name = "txtCustValue"
        Me.txtCustValue.Size = New System.Drawing.Size(336, 20)
        Me.txtCustValue.TabIndex = 4
        '
        'txtCustName
        '
        Me.txtCustName.Location = New System.Drawing.Point(64, 248)
        Me.txtCustName.Name = "txtCustName"
        Me.txtCustName.Size = New System.Drawing.Size(144, 20)
        Me.txtCustName.TabIndex = 2
        '
        'cmdRemove
        '
        Me.cmdRemove.Enabled = False
        Me.cmdRemove.Location = New System.Drawing.Point(328, 200)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.Size = New System.Drawing.Size(80, 24)
        Me.cmdRemove.TabIndex = 6
        Me.cmdRemove.Text = "Remove"
        '
        'cmdAdd
        '
        Me.cmdAdd.Location = New System.Drawing.Point(240, 200)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(80, 24)
        Me.cmdAdd.TabIndex = 5
        Me.cmdAdd.Text = "Add"
        '
        'CustListView
        '
        Me.CustListView.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.CustNameCol, Me.CustValueCol, Me.CustTypeCol, Me.CustCountCol})
        Me.CustListView.FullRowSelect = True
        Me.CustListView.GridLines = True
        Me.CustListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.CustListView.HideSelection = False
        Me.CustListView.Location = New System.Drawing.Point(16, 32)
        Me.CustListView.MultiSelect = False
        Me.CustListView.Name = "CustListView"
        Me.CustListView.Size = New System.Drawing.Size(420, 160)
        Me.CustListView.TabIndex = 1
        Me.CustListView.UseCompatibleStateImageBehavior = False
        Me.CustListView.View = System.Windows.Forms.View.Details
        '
        'CustNameCol
        '
        Me.CustNameCol.Text = "Name"
        Me.CustNameCol.Width = 120
        '
        'CustValueCol
        '
        Me.CustValueCol.Text = "Value"
        Me.CustValueCol.Width = 186
        '
        'CustTypeCol
        '
        Me.CustTypeCol.Text = "Type"
        Me.CustTypeCol.Width = 67
        '
        'CustCountCol
        '
        Me.CustCountCol.Text = "Count"
        Me.CustCountCol.Width = 43
        '
        'lbCustIntro
        '
        Me.lbCustIntro.Location = New System.Drawing.Point(16, 8)
        Me.lbCustIntro.Name = "lbCustIntro"
        Me.lbCustIntro.Size = New System.Drawing.Size(344, 16)
        Me.lbCustIntro.TabIndex = 8
        Me.lbCustIntro.Text = "Custom Document Properties:"
        '
        'FilePropDemo
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(479, 416)
        Me.Controls.Add(Me.cmdOpen)
        Me.Controls.Add(Me.PropTabs)
        Me.Controls.Add(Me.lbFileName)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "FilePropDemo"
        Me.Text = "Custom Properties Manager"
        Me.PropTabs.ResumeLayout(False)
        Me.CustomTab.ResumeLayout(False)
        Me.CustomTab.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region

    ' OpenDocumentProperties:

    Public Function OpenDocumentProperties() As Boolean

        Dim oDialog As OpenFileDialog
        Dim sFile As String
        Dim fOpenReadOnly As Boolean

        ' Ask the user for an Office OLE Structure Storage file to read
        ' the document properties from. You can also select any common file
        ' on a Win2K/WinXP NTFS drive, or a non-Office OLE file...
        oDialog = New OpenFileDialog
        oDialog.Filter = c_strFilter
        oDialog.FilterIndex = 1
        oDialog.ShowReadOnly = True
        oDialog.Multiselect = True
        oDialog.CheckFileExists = True
        If oDialog.ShowDialog() = DialogResult.Cancel Then
            ' Nothing to do if user cancels...
            Return False
        End If

        CloseCurrentDocument(True)
        filenames = New List(Of String)

        Dim ToolTip1 = New System.Windows.Forms.ToolTip()
        Dim fileList = ""

        For Each sFile In oDialog.FileNames

            fOpenReadOnly = oDialog.ReadOnlyChecked
            filenames.Add(sFile)
            fileList += Path.GetFileName(sFile) & vbCrLf

            lbFileName.Text = filenames.Count & " file(s) are currently open"

            'Dim document As WordprocessingDocument
            Using document = WordprocessingDocument.Open(sFile, True)
                Dim oCustProp As CustomDocumentProperty
                For Each oCustProp In document.CustomFilePropertiesPart.Properties
                    If oCustProp.Name Is Nothing Then
                        oCustProp.Remove()
                    End If
                    Dim vType As PropertyTypes
                    Dim asd As String
                    Try
                        'MsgBox(oCustProp.VTBool.InnerText)
                        asd = oCustProp.VTBool.InnerText
                        vType = PropertyTypes.YesNo
                    Catch
                    End Try
                    Try
                        'MsgBox(oCustProp.VTLPWSTR.InnerText)
                        asd = oCustProp.VTLPWSTR.InnerText
                        vType = PropertyTypes.Text
                    Catch
                    End Try
                    Try
                        'MsgBox(oCustProp.VTFileTime.InnerText)
                        asd = oCustProp.VTFileTime.InnerText
                        vType = PropertyTypes.DateTime
                    Catch
                    End Try
                    Try
                        'MsgBox(oCustProp.VTFloat.InnerText)
                        asd = oCustProp.VTFloat.InnerText
                        vType = PropertyTypes.NumberDouble
                    Catch
                    End Try
                    Try
                        'MsgBox(oCustProp.VTInt32.InnerText)
                        asd = oCustProp.VTInt32.InnerText
                        vType = PropertyTypes.NumberInteger
                    Catch
                    End Try
                    CustCheckMultipleLv(oCustProp.Name, CStr(oCustProp.InnerText), vType)
                    'MsgBox(oCustProp.Name)

                Next oCustProp

            End Using


            ' Enable/Disable text items if file is open read only...
            EnableItems(Not m_fOpenedReadOnly)

            ' The operation was successful.
            OpenDocumentProperties = True
            PropTabs.Enabled = True
        Next sFile

        ToolTip1.SetToolTip(lbFileName, fileList)
        ToolTip1.AutoPopDelay = 10000
    End Function


    ' CloseCurrentDocument:
    '  Closed the open document and clears the dialog. Can prompt the user
    '  to save the changes made if open read-write mode.
    Private Sub CloseCurrentDocument(ByVal bPromptToSaveIfDirty As Boolean)
        ' Reset all the control values to default...
        ClearControls()
        PropTabs.Enabled = False
    End Sub

    ' EnableItems:
    '   Helper function to enable/disable controls with respect to
    '   read-only mode for the current file...
    Private Sub EnableItems(ByVal bEnable As Boolean)
        txtCustValue.Enabled = bEnable : cboxCustType.Enabled = bEnable
        cmdAdd.Enabled = bEnable
    End Sub

    ' ClearControls:
    '  Helper function to restore dialog controls to "blank" slate...
    Private Sub ClearControls()
        lbFileName.Text = "[Click Open button to read properties from file...]"
        txtCustName.Text = "" : txtCustValue.Text = ""
        cboxCustType.SelectedIndex = 0
        CustListView.Items.Clear()
        cmdRemove.Enabled = False
    End Sub

    ' Check labels
    ' and write it
    Private Sub CheckMultiple(ByRef labelText As String, ByVal docuProp As String, Optional ByVal empty As String = ""
                                )
        If labelText = empty Then
            labelText = docuProp
        Else
            If docuProp <> labelText Then
                labelText = "<Multiple Values>"
            End If
        End If
    End Sub



    Private Sub CustCheckMultipleLv(ByVal ColName As String, ByVal PropName As String, Optional ByVal vType As PropertyTypes = PropertyTypes.Unknown)
        Dim lvItem As ListViewItem
        Dim tmpLvItem As ListViewItem
        lvItem = GetLvItemForProperty(ColName, PropName, vType)
        tmpLvItem = Nothing

        'Check item is not in ListView yet
        For Each sLvItem In CustListView.Items
            If sLvItem.SubItems.Item(0).Text = ColName Then
                tmpLvItem = sLvItem
            End If
        Next

        If tmpLvItem Is Nothing Then
            CustListView.Items.Add(lvItem)
        Else
            CheckMultiple(tmpLvItem.SubItems.Item(1).Text, lvItem.SubItems.Item(1).Text)
            'MsgBox(tmpLvItem.SubItems.Item(2).Text)
            tmpLvItem.SubItems.Item(3).Text = tmpLvItem.SubItems.Item(3).Text + 1
        End If
    End Sub

    ' GetLvItemForProperty:
    '  Helper function to take name and value and return ListViewItem 
    '  which can be added to the listview for the propset being displayed.
    Private Function GetLvItemForProperty(ByVal sName As String, ByVal sValue As String, Optional ByVal vType As PropertyTypes = PropertyTypes.Unknown)
        Dim lvItem As New ListViewItem
        Dim sTypeName As String
        lvItem.Text = sName
        lvItem.SubItems.Add(sValue)
        Select Case vType
            Case PropertyTypes.Text
                sTypeName = "String"
            Case PropertyTypes.NumberDouble
                sTypeName = "Double"
            Case PropertyTypes.NumberInteger
                sTypeName = "Int"
            Case PropertyTypes.DateTime
                sTypeName = "Date"
            Case PropertyTypes.YesNo
                sTypeName = "Bool"
            Case Else
                sTypeName = "Unknown"
        End Select
        lvItem.SubItems.Add(sTypeName)
        lvItem.SubItems.Add(1)
        GetLvItemForProperty = lvItem
    End Function

    ' GetIconImageFromDisp:
    '  Helper function to take PictureDisp and create GDI+ Bitmap Image for 
    '  the file icon to display in WinForm PictureBox...
    Private Function GetIconImageFromDisp(ByVal oDispPicture As Object) As System.Drawing.Image
        Dim iType As Integer
        Dim iHandle As Integer
        Dim args() As Object
        Try
            ' Confirm that object contains an ICON picture...
            iType = CLng(oDispPicture.GetType.InvokeMember("Type", Reflection.BindingFlags.GetProperty, Nothing, oDispPicture, args))
            If iType = 3 Then ' If So, ask for the handle...
                iHandle = oDispPicture.GetType.InvokeMember("Handle", Reflection.BindingFlags.GetProperty, Nothing, oDispPicture, args)
                ' Create the Drawing.Bitmap object from the ICON handle...
                GetIconImageFromDisp = System.Drawing.Bitmap.FromHicon(New System.IntPtr(iHandle))
            End If
        Catch ex As Exception
            ' Return Nothing if exception thrown..
        End Try
    End Function

    ' cmdOpen_Click: Handler for Open Button
    Private Sub cmdOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpen.Click
        OpenDocumentProperties()
        'CloseCurrentDocument(True)
        'If Not OpenDocumentProperties() Then
        'CloseCurrentDocument(False)
        'End If
    End Sub

    ' cmdAdd_Click: Handler for Custom Property Add Button
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim sName As String, sValueText As String
        Dim vValue As Object

        sName = txtCustName.Text
        sValueText = txtCustValue.Text

        ' We can't add a custom property unless we have a valid name
        ' and value, so we do a quick check now to avoid error later on.
        If ((sName = "") Or (sValueText = "")) Then
            Beep() : txtCustName.Select()
            Exit Sub
        End If
        Dim type As PropertyTypes
        Try
            Select Case cboxCustType.SelectedIndex + 1
                Case 2
                    vValue = CInt(sValueText)
                    type = PropertyTypes.NumberInteger
                Case 3
                    vValue = CDbl(sValueText)
                    type = PropertyTypes.NumberDouble
                Case 4
                    vValue = CBool(sValueText)
                    type = PropertyTypes.YesNo
                Case 5
                    vValue = CDate(sValueText)
                    type = PropertyTypes.DateTime
                Case Else
                    vValue = sValueText
                    type = PropertyTypes.Text
            End Select
        Catch ex As Exception
            MsgBox("Invalid conversion of" & sValueText & vbCrLf & "Error: " & ex.Message, MsgBoxStyle.Critical)
            Return
        End Try

        For Each sFile In filenames
            Try
                SetCustomProperty(sFile, sName, vValue, type)
            Catch ex As Exception
                MsgBox("The item could not be added!" & vbCrLf & "Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        Next

        ' If that succeeded, add the item to the list view...
        Dim lvItem = GetLvItemForProperty(sName, sValueText, type)
        lvItem.Subitems.Item(3).Text = filenames.Count

        Dim tmpLvItem = CustListView.FindItemWithText(sName)
        If Not tmpLvItem Is Nothing Then
            tmpLvItem.Remove()
        End If

        CustListView.Items.Add(lvItem)
        txtCustName.Text = "" : txtCustValue.Text = ""

    End Sub

    ' cmdRemove_Click: Handler for Custom Property Remove Button
    Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove.Click
        Dim iRemoveCnt As Integer, i As Integer
        Dim sName As String

        ' Get the count of selected items (normally this is 1)...
        iRemoveCnt = CustListView.SelectedItems.Count
        If iRemoveCnt < 1 Then ' This should not happen, but just in case...
            MsgBox("There is no selected item to remove!", MsgBoxStyle.Critical)
            Exit Sub
        End If

        ' For each fo the selected items, remove them from the document.
        ' We loop backwards to not change the list as we also want to 
        ' remove the listview item when the property is removed...
        For i = iRemoveCnt - 1 To 0 Step -1
            sName = CustListView.SelectedItems(i).Text
            Try
                CustListView.SelectedItems(i).Remove()
                For Each sFile In filenames
                    RemoveCustomProperty(sFile, sName)
                Next
            Catch ex As Exception
                MsgBox("Unable to remove '" & sName & _
                    vbCrLf & "Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        Next

        cmdRemove.Enabled = False
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustListView.DoubleClick

        txtCustName.Text = CustListView.SelectedItems(0).SubItems.Item(0).Text
        If CustListView.SelectedItems(0).SubItems.Item(1).Text = "<Multiple Values>" Then
            txtCustValue.Text = ""
        Else
            txtCustValue.Text = CustListView.SelectedItems(0).SubItems.Item(1).Text
        End If
        Select Case CustListView.SelectedItems(0).SubItems.Item(2).Text
            Case "String"
                cboxCustType.SelectedIndex = 0
            Case "Int"
                cboxCustType.SelectedIndex = 1
            Case "Double"
                cboxCustType.SelectedIndex = 2
            Case "Bool"
                cboxCustType.SelectedIndex = 3
            Case "Date"
                cboxCustType.SelectedIndex = 4
            Case Else
                cboxCustType.SelectedIndex = 0
        End Select
    End Sub


    ' FilePropDemo_Closing: Cleanup before exit and ask user to save changes if needed...
    Private Sub FilePropDemo_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        CloseCurrentDocument(True)
    End Sub

    ' CustListView_SelectedIndexChanged: Enable Remove button if item selected.
    Private Sub CustListView_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustListView.SelectedIndexChanged
        If (Not m_fOpenedReadOnly) Then
            cmdRemove.Enabled = CustListView.SelectedItems.Count > 0
        End If
    End Sub

    Public Enum PropertyTypes
        YesNo
        Text
        DateTime
        NumberInteger
        NumberDouble
        Unknown
    End Enum

    Public Function SetCustomProperty( _
        ByVal fileName As String,
        ByVal propertyName As String, _
        ByVal propertyValue As Object,
        ByVal propertyType As PropertyTypes) As String

        ' Given a document name, a property name/value, and the property type, 
        ' add a custom property to a document. The method returns the original 
        ' value, if it existed.

        Dim returnValue As String = Nothing

        Dim newProp As New CustomDocumentProperty
        Dim propSet As Boolean = False

        ' Calculate the correct type:
        Select Case propertyType

            Case PropertyTypes.DateTime
                ' Make sure you were passed a real date, 
                ' and if so, format in the correct way. 
                ' The date/time value passed in should 
                ' represent a UTC date/time.
                If TypeOf (propertyValue) Is DateTime Then
                    newProp.VTFileTime = _
                        New VTFileTime(String.Format("{0:s}Z",
                            Convert.ToDateTime(propertyValue)))
                    propSet = True
                End If

            Case PropertyTypes.NumberInteger
                If TypeOf (propertyValue) Is Integer Then
                    newProp.VTInt32 = New VTInt32(propertyValue.ToString())
                    propSet = True
                End If

            Case PropertyTypes.NumberDouble
                If TypeOf propertyValue Is Double Then
                    newProp.VTFloat = New VTFloat(propertyValue.ToString())
                    propSet = True
                End If

            Case PropertyTypes.Text
                newProp.VTLPWSTR = New VTLPWSTR(propertyValue.ToString())
                propSet = True

            Case PropertyTypes.YesNo
                If TypeOf propertyValue Is Boolean Then
                    ' Must be lowercase.
                    newProp.VTBool = _
                      New VTBool(Convert.ToBoolean(propertyValue).ToString().ToLower())
                    propSet = True
                End If
        End Select

        If Not propSet Then
            ' If the code was not able to convert the 
            ' property to a valid value, throw an exception.
            Throw New InvalidDataException("propertyValue")
        End If

        ' Now that you have handled the parameters, start
        ' working on the document.
        newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
        newProp.Name = propertyName

        Using document = WordprocessingDocument.Open(fileName, True)
            Dim customProps = document.CustomFilePropertiesPart
            If customProps Is Nothing Then
                ' No custom properties? Add the part, and the
                ' collection of properties now.
                customProps = document.AddCustomFilePropertiesPart
                customProps.Properties = New Properties
            End If

            Dim props = customProps.Properties
            If props IsNot Nothing Then
                ' This will trigger an exception is the property's Name property 
                ' is null, but if that happens, the property is damaged, and 
                ' probably should raise an exception.
                Dim prop = props.Where(Function(p) CType(p, CustomDocumentProperty).
                          Name.Value = propertyName).FirstOrDefault()
                ' Does the property exist? If so, get the return value, 
                ' and then delete the property.
                If prop IsNot Nothing Then
                    returnValue = prop.InnerText
                    prop.Remove()
                End If

                ' Append the new property, and 
                ' fix up all the property ID values. 
                ' The PropertyId value must start at 2.
                props.AppendChild(newProp)
                Dim pid As Integer = 2
                For Each item As CustomDocumentProperty In props
                    item.PropertyId = pid
                    pid += 1
                Next
                props.Save()
            End If
        End Using

        Return returnValue

    End Function

    Public Function RemoveCustomProperty( _
        ByVal fileName As String,
        ByVal propertyName As String
        ) As Boolean

        ' Given a document name, a property name/value, and the property type, 
        ' add a custom property to a document. The method returns the original 
        ' value, if it existed.

        Dim newProp As New CustomDocumentProperty
        Dim propSet As Boolean = False

        Using document = WordprocessingDocument.Open(fileName, True)
            Dim customProps = document.CustomFilePropertiesPart
            If customProps Is Nothing Then
                ' No custom properties
                Return True
            End If

            Dim props = customProps.Properties
            If props IsNot Nothing Then
                ' This will trigger an exception if the property's Name property 
                ' is null, but if that happens, the property is damaged, and 
                ' probably should raise an exception.
                Dim prop = props.Where(Function(p) CType(p, CustomDocumentProperty).
                          Name.Value = propertyName).FirstOrDefault()
                ' Does the property exist? If so, get the return value, 
                ' and then delete the property.
                If prop Is Nothing Then
                    'Do nothing, custom property is not there
                    Return True
                End If

                ' Remove the property, and 
                ' fix up all the property ID values. 
                ' The PropertyId value must start at 2.
                prop.Remove()
                Dim pid As Integer = 2
                For Each item As CustomDocumentProperty In props
                    item.PropertyId = pid
                    pid += 1
                Next
                props.Save()
            End If
        End Using

        Return True

    End Function


    Public Function GetCustomProperty( _
       ByVal fileName As String,
       ByVal propertyName As String
       ) As CustomDocumentProperty

        ' Given a document name, a property name/value, and the property type, 
        ' get a custom property to a document. The method returns the original 
        ' value, if it existed.

        Using document = WordprocessingDocument.Open(fileName, True)
            Dim customProps = document.CustomFilePropertiesPart
            If customProps Is Nothing Then
                ' No custom properties
                Return Nothing
            End If

            Dim props = customProps.Properties
            If props IsNot Nothing Then
                ' This will trigger an exception i the property's Name property 
                ' is null, but if that happens, the property is damaged, and 
                ' probably should raise an exception.
                Dim prop = props.Where(Function(p) CType(p, CustomDocumentProperty).
                          Name.Value = propertyName).FirstOrDefault()
                If prop Is Nothing Then
                    'Do nothing, prop is not there
                    Return Nothing
                End If
                Return prop
            End If
        End Using
        Return Nothing
    End Function

    Friend WithEvents PropTabs As System.Windows.Forms.TabControl
    Friend WithEvents CustomTab As System.Windows.Forms.TabPage
    Friend WithEvents cboxCustType As System.Windows.Forms.ComboBox
    Friend WithEvents lbCustType As System.Windows.Forms.Label
    Friend WithEvents lbCustValue As System.Windows.Forms.Label
    Friend WithEvents lbCustName As System.Windows.Forms.Label
    Friend WithEvents lbCustNote As System.Windows.Forms.Label
    Friend WithEvents txtCustValue As System.Windows.Forms.TextBox
    Friend WithEvents txtCustName As System.Windows.Forms.TextBox
    Friend WithEvents cmdRemove As System.Windows.Forms.Button
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents CustListView As System.Windows.Forms.ListView
    Friend WithEvents CustNameCol As System.Windows.Forms.ColumnHeader
    Friend WithEvents CustValueCol As System.Windows.Forms.ColumnHeader
    Friend WithEvents CustTypeCol As System.Windows.Forms.ColumnHeader
    Friend WithEvents CustCountCol As System.Windows.Forms.ColumnHeader
    Friend WithEvents lbCustIntro As System.Windows.Forms.Label

End Class

Imports System
Imports System.Diagnostics
Imports System.ComponentModel
Imports Microsoft.Win32
Imports System.Text
Imports System.Management
Imports Microsoft.Office

Public Class Form1
    Dim dt As New DataTable
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dt.Columns.Add("Program")
        'GetInstallApplications()
        'Dim p() As Process
        'p = Process.GetProcesses()
        'For Each Pro As AppInfo In GetInstalledApps()
        '    If Pro.UnInstallPath IsNot Nothing Then
        '        Dim LI As New ListViewItem
        '        LI.SubItems.Add(Pro.UnInstallPath.Replace("""", ""))
        '        LI.SubItems.Add(Pro.SilentUnInstallPath)
        '        ListView1.Items.Add(LI.Name)
        '    End If
        ' ListBox1.Items.AddRange(New Object() {Pro.Id, Pro.ProcessName})
        ' ListBox1.Items.Add(Pro.ProcessName & " # " & Pro.MainWindowTitle & " # " & Pro.Id)
        'Next
        For Each AI As AppInfo In GetInstalledApps()
            If AI.UnInstallPath IsNot Nothing Then
                Dim LI As New ListViewItem
                LI.Text = AI.Name
                LI.SubItems.Add(AI.UnInstallPath.Replace("""", ""))
                LI.SubItems.Add(AI.SilentUnInstallPath)
                ListView1.Items.Add(LI)
                ' dt.Rows.Add(LI)
            End If
        Next
        For i As Integer = 0 To ListView1.Items.Count - 1
            dt.Rows.Add(ListView1.Items(i).SubItems(0).Text)
        Next
        DataGridView1.DataSource = dt
    End Sub
    'Private Function GetInstallApplications() As String
    '    Dim key As RegistryKey
    '    Dim st As New StringBuilder
    '    key = Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Uninstall\")
    '    For Each subkey As String In key.GetSubKeyNames
    '        st.AppendLine(subkey)
    '        dt.Rows.Add(subkey)
    '    Next
    '    Return st.ToString
    'End Function
    Structure AppInfo
        Dim Name As String
        Dim UnInstallPath As String
        Dim SilentUnInstallPath As String
    End Structure
    Private Function GetInstalledApps() As List(Of AppInfo)
        Dim DestKey As String = "Software\Microsoft\Windows\CurrentVersion\Uninstall\"
        Dim iList As New List(Of AppInfo)
        For Each app As String In Registry.LocalMachine.OpenSubKey(DestKey).GetSubKeyNames
            iList.Add(New AppInfo With {.Name = Registry.LocalMachine.OpenSubKey(DestKey & app & "\").GetValue("DisplayName"), _
                        .UnInstallPath = Registry.LocalMachine.OpenSubKey(DestKey & app & "\").GetValue("UninstallString")})
        Next
        Return iList
    End Function

    Private Sub ExportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToolStripMenuItem.Click
        Dim xlApp As Interop.Excel.Application
        Dim xlWorkBook As Interop.Excel.Workbook
        Dim xlWorkSheet As Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer

        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")


        For i = 0 To DataGridView1.RowCount - 2
            For j = 0 To DataGridView1.ColumnCount - 1
                For k As Integer = 1 To DataGridView1.Columns.Count
                    xlWorkSheet.Cells(1, k) = DataGridView1.Columns(k - 1).HeaderText
                    xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()
                Next
            Next
        Next

        xlWorkSheet.SaveAs("D:\" & Environment.UserName & "vbexcel.xlsx")
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        MsgBox("You can find the file D:\vbexcel.xlsx")
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class

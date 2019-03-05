Option Strict Off
#Region "Imports"
Imports System.Collections.Generic
Imports System.Text
Imports JDA.Intactix.Automation
Imports JDA.Intactix.Automation.Space
Imports System.Xml
Imports System.IO
#End Region

Public Class SpaceMenuClass
    Inherits Script
    Public Sub New()
        MyBase.New()
    End Sub

   Public Sub New(ByVal nSpaceOrFloor As Integer)
        MyBase.New(nSpaceOrFloor)
    End Sub

    Public Overrides Sub Run(Optional isSilentMode As Boolean = False)
        SetActiveProjectDB()
        SetActivePlanogram()
        DeletePositionXML()
        SaveVersionToDatabase()
    End Sub

    Dim mySegment
    Public Sub SetActiveSegment()
        'To set the active Segment
        Try
            mySegment = InputBox("Enter the Segment Number from 0-N")
            SpacePlanning.DeselectAllObjects()
            SpacePlanning.SelectSegment(mySegment)
            'MsgBox("Segment " & mySegment & " has been set")
        Catch
            MsgBox("Enter a valid Segment Number")
            Return
        End Try
    End Sub

    Dim myFixture
    Public Sub SetActiveFixture()
        'To set the active Fixtrue
        Try
            myFixture = InputBox("Enter the Fixture Number from 0-N")
            SpacePlanning.DeselectAllObjects()
            SpacePlanning.SelectFixture(myFixture)
            'MsgBox("Fixture " & myFixture & " has been set")
        Catch
            MsgBox("Enter a valid Fixture Number")
            Return
        End Try
    End Sub
    Public Sub SetActiveProjectFile()
        'To open the Projetc via File Browser by Storing the Path in to a Variable
        Dim oExec = CreateObject("WScript.Shell").Exec("mshta.exe ""about:" & "<" & "input type=file id=FILE>" & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>""")
        Dim Tst = oExec.StdOut.ReadAll
        Tst = Replace(Tst, vbCrLf, "")
        SpacePlanning.OpenProjectFile(Tst)
    End Sub

    Dim myPlanogram
    Public Sub SetActivePlanogram()
        'To set the active Planogram
        Try
            myPlanogram = InputBox("Enter the Planogram Number from 0-N")
            SpacePlanning.DeselectAllObjects()
            SpacePlanning.SetActivePlanogram(myPlanogram)
            'MsgBox("Planogram " & myPlanogram & " has been set")
        Catch
            MsgBox("Enter a Valid Planogram Number")
            Return
        End Try
    End Sub
    Public Sub SetActiveProjectDB()
        'To select the Database using DB Key ID
        Try
            Dim dbKey
            SpacePlanning.SelectDatabase("CKB | ACCECCTRN\SQLEXPRESS01 | ")
            dbKey = InputBox("Enter the DB Key")
            SpacePlanning.OpenProjectFromDatabase(dbKey)
            'MsgBox("DB " & dbKey & "has been set")
        Catch ex As Exception
            MsgBox("Enter the right Database Key")
            Return
        End Try
    End Sub
    Public Sub DeletePositionOneMany()
        'To retrieve the Products in the Shelf and to delete them.
        Try
            Dim myProduct
            Dim myChoice
            myProduct = InputBox("Enter the Product ID to be found")
            Dim Prod = SpacePlanning.FindProduct("myProduct")
            If Prod IsNot Nothing Then
                MsgBox("Enter the right Product ID")
            Else
                MsgBox("Product Found")
                myChoice = InputBox("Delete the Product from Active Planogram or All" & vbCrLf & "Enter 1 to select active Planogram and 2 for All Planograms")
                If (myChoice = 1) Then
                    SpacePlanning.DeselectAllObjects()
                    SpacePlanning.SetActivePlanogram(0)
                    SpacePlanning.DeletePosition("ID=" + myProduct + ";UPC=" + myProduct, "1")
                    MsgBox("Product Deleted Successfully")
                    Return
                ElseIf (myChoice = 2) Then
                    Dim SpProj As Space.Project
                    Dim i As Integer = -1
                    SpProj = SpacePlanning.OpenProjectFile("C:\Program Files (x86)\JDA\Intactix\Samples\Space Planning Planograms\Sample - Ice Cream.psa")
                    For Each plano As Space.Planogram In SpProj.Planograms()
                        i = i + 1
                        SpacePlanning.SetActivePlanogram(i)
                        SpacePlanning.DeletePosition("ID=" + myProduct + ";UPC=" + myProduct, "1")
                    Next
                    MsgBox("All the Products have been deleted successfully")
                Else
                    MsgBox("Enter a valid choice")
                    Return
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return
        End Try
    End Sub

Public Sub DeletePositionXML()
        Dim list As List(Of String) = New List(Of String)
        Dim xml As XmlDocument = New XmlDocument
        Dim i As Integer = 0
        Dim oExec = CreateObject("WScript.Shell").Exec("mshta.exe ""about:" & "<" & "input type=file id=FILE>" & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>""")
        Dim Tst = oExec.StdOut.ReadAll
        Tst = Replace(Tst, vbCrLf, "")
        xml.Load(Tst)
        Dim myProduct
        SpacePlanning.DeselectAllObjects()
        SpacePlanning.SetActivePlanogram(0)
        Try
            For Each ioObject As XmlNode In xml.SelectNodes("//productid")
                list.Add(ioObject.InnerXml)
                myProduct = list.Item(i)
                SpacePlanning.DeletePosition("ID=" + myProduct + ";UPC=" + myProduct)
                i = i + 1
            Next
            MsgBox("Items Deleted Successfully")
        Catch exception As Exception
            MsgBox("Sorry")
        End Try
    End Sub
    Public Sub SaveVersionToDatabase()
        Try
            Dim Answer As String
            Dim MyNote As String
            'Place your text here
            MyNote = "Save the Planogrma as Another Version?"
            'Display MessageBox
            Answer = MsgBox(MyNote, CType(vbQuestion + vbYesNo, MsgBoxStyle), "Save")
            If Answer = vbNo Then
                Return
            Else
                SpacePlanning.SaveAsPlanogramVersionToDatabase("Unassigned", 3, 3, True, True)
                MsgBox("Saved Successfully")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
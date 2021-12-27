'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.DirectoryServices.AccountManagement
Imports System.Environment
Imports System.IO
Imports System.Linq

'This module contains this program's core procedures.
Public Module UserGroupsModule
   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim Environment As PrincipalContext = Nothing
         Dim PathO As String = Nothing
         Dim UserSearchName As String = Nothing

         With My.Application.Info
            Console.WriteLine($"{ .Title} v{ .Version} - by: { .CompanyName}")
            Console.WriteLine()
         End With

         Select Case GetChoice("(D = domain, M = machine, Q = quit) Choice: ", "DMQdmq").ToLower()
            Case "d"
               Environment = New PrincipalContext(ContextType.Domain)
            Case "m"
               Environment = New PrincipalContext(ContextType.Machine)
            Case "q"
               Exit Sub
         End Select

         Console.WriteLine()
         Console.Write("User name: ")
         UserSearchName = Console.ReadLine()

         If Not UserSearchName = Nothing Then
            Console.WriteLine()
            Console.Write($"Specify path (default: ""{Directory.GetCurrentDirectory()}""): ")
            PathO = Console.ReadLine()
            Console.WriteLine()

            SearchForUsers(Environment, UserSearchName, PathO)

            Console.WriteLine()
            Console.WriteLine("Done. - Press Enter to continue...")
            Console.ReadLine()
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays a prompt and requests the user to make a choice.
   Private Function GetChoice(Prompt As String, Choices As String) As String
      Try
         Dim Choice As String = Nothing

         Console.Write(Prompt)
         Do
            Choice = Console.ReadKey(intercept:=True).KeyChar.ToString()
         Loop Until Choices.Contains(Choice)
         Console.WriteLine(Choice)

         Return Choice
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns a list of groups for the specified user.
   Private Function GetGroups(User As UserPrincipal) As List(Of Principal)
      Try
         Return New List(Of Principal)(User.GetGroups)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns a list of all users in the specified environment.
   Private Function GetUsers(Environment As PrincipalContext) As List(Of Principal)
      Try
         Return New List(Of Principal)((New PrincipalSearcher() With {.QueryFilter = New UserPrincipal(Environment) With {.Name = "*"}}).FindAll())
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure handles any errors that occor.
   Private Sub HandleError(ExceptionO As Exception)
      Try
         Console.ForegroundColor = ConsoleColor.Red
         Console.WriteLine()
         Console.WriteLine($"Error: {ExceptionO.Message}")
         Console.WriteLine("Press Enter to continue...")
         Console.ReadLine()
         Console.ResetColor()
      Catch
         [Exit](0)
      End Try
   End Sub

   'This procedure saves the specified user's group data.
   Private Sub SaveGroupData(UserName As String, GroupList As List(Of Principal), PathO As String)
      Try
         Dim GroupTable As New DataTable(UserName)

         GroupTable.Columns.Add("Group")
         GroupTable.Columns.Add("Discription")

         GroupList.ForEach(Sub(Group As Principal) GroupTable.Rows.Add(Group.Name, Group.Description))

         GroupTable.WriteXml(Path.Combine(PathO, UserName & ".xml"))
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure searches for the specified user.
   Private Sub SearchForUsers(Environment As PrincipalContext, UserSearchName As String, PathO As String)
      Try
         Console.WriteLine($"Users found in: {Environment.ConnectedServer}")

         For Each FoundUser As UserPrincipal In From User As Principal In GetUsers(Environment) Where User.Name.ToLower().Contains(UserSearchName.ToLower())
            Console.WriteLine(FoundUser.ToString())
            SaveGroupData(FoundUser.Name, GetGroups(FoundUser), PathO)
         Next FoundUser
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub
End Module

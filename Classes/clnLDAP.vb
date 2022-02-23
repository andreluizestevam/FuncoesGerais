Imports System
Imports System.DirectoryServices

Public Class clnLDAP

    Public Shared Function Authentication(ByVal path As String, ByVal user As String, ByVal password As String) As String
        Try
            Dim oAD As DirectoryEntry = New DirectoryEntry("LDAP://" + path, user, password)
            'Caso não encontra o usuário no AD, volta a mensagem de erro.
            oAD.SchemaClassName.ToString()
            Return "TRUE"

        Catch ex As Exception
            Return ex.Message
        End Try

    End Function


    Public Shared Function UserListAD(ByVal path As String, ByVal user As String, ByVal password As String) As DataTable
        Dim table As DataTable = New DataTable("Resultados")
        table.Columns.Add("Nome")
        table.Columns.Add("Usuario")
        table.Columns.Add("Email")
        Dim row As DataRow
        Dim deRoot As DirectoryEntry = New DirectoryEntry("LDAP://" + path, user, password)
        Dim deSrch As DirectorySearcher = New DirectorySearcher(deRoot, "(&(objectClass=user)(objectCategory=person))")
        deSrch.PropertiesToLoad.Add("cn")
        deSrch.PropertiesToLoad.Add("userPrincipalName")
        deSrch.PropertiesToLoad.Add("sAMAccountName")
        deSrch.PropertiesToLoad.Add("mail")
        deSrch.PropertiesToLoad.Add("password")
        deSrch.Sort.PropertyName = "sAMAccountName"
        Dim oRes As SearchResult
        For Each oRes In deSrch.FindAll()
            row = table.NewRow()
            row("Nome") = oRes.Properties("cn")(0).ToString()
            row("usuario") = oRes.Properties("sAMAccountName")(0).ToString()
            If oRes.Properties.Contains("mail") Then
                row("Email") = oRes.Properties("mail")(0).ToString()
            End If

            table.Rows.Add(row)
        Next

        Return table

    End Function

    Public Shared Function ChangePassword(ByVal path As String, ByVal user As String, ByVal password_old As String, ByVal password_new As String) As String
        Try
            Dim de As New DirectoryEntry("LDAP://" + path, user, password_old)
            Dim deSearch As DirectorySearcher = New DirectorySearcher
            deSearch.SearchRoot = de
            deSearch.Filter = "(&(objectClass=user)(sAMAccountName=" + user + "))"
            deSearch.SearchScope = SearchScope.Subtree
            Dim results As SearchResult = deSearch.FindOne()
            If Not (results Is Nothing) Then
                de = New DirectoryEntry(results.Path, user, password_old, AuthenticationTypes.Secure)
                de.Invoke("ChangePassword", New Object() {password_old, password_new})
                de.CommitChanges()
                Return "TRUE"
            End If
        Catch ex As Exception
            Return ex.InnerException.Message
        End Try

    End Function

    Public Shared Function ReturnUserAD(ByVal path As String, ByVal user As String) As DataTable
        Dim table As DataTable = New DataTable("Resultados")
        table.Columns.Add("Nome")
        table.Columns.Add("Usuario")
        table.Columns.Add("Email")
        Dim row As DataRow
        Dim deRoot As DirectoryEntry = New DirectoryEntry("LDAP://" + path, "uniaoquimica\sepuser", "P@ssw0rd")
        Dim deSrch As DirectorySearcher = New DirectorySearcher(deRoot, "(&(objectClass=user)(objectCategory=person))")
        deSrch.PropertiesToLoad.Add("cn")
        deSrch.PropertiesToLoad.Add("userPrincipalName")
        deSrch.PropertiesToLoad.Add("sAMAccountName")
        deSrch.PropertiesToLoad.Add("mail")
        deSrch.PropertiesToLoad.Add("password")
        deSrch.Sort.PropertyName = "sAMAccountName"
        Dim oRes As SearchResult
        For Each oRes In deSrch.FindAll
            row = table.NewRow()
            row("Nome") = oRes.Properties("cn")(0).ToString()
            row("usuario") = oRes.Properties("sAMAccountName")(0).ToString()
            If oRes.Properties.Contains("mail") Then
                row("Email") = oRes.Properties("mail")(0).ToString()
            End If

            table.Rows.Add(row)
        Next

        Dim dvUser As DataView = table.DefaultView
        dvUser.RowFilter = "usuario = '" & user & "'"
        Dim dtUser As DataTable = dvUser.ToTable("Table")

        Return dtUser

    End Function

    
End Class

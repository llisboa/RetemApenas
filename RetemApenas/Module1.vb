Module Module1
    Function TemporaryDir(Optional ByVal Dir As String = "") As String
        Return System.IO.Path.GetTempPath
    End Function


    Public Function DosShell(ByVal Comando As String, ByVal Argumento As String, ByVal Diretorio As String, ByVal Entrada As String, ByRef Erros As String, Optional ByVal Usuario As String = "", Optional ByVal Senha As String = "", Optional ByVal Dominio As String = "", Optional ByVal EsperaSegs As Integer = 30) As String
        Dim Result As String = ""
        Dim Proc As System.Diagnostics.Process = Nothing
        Dim StdIn As System.IO.StreamWriter = Nothing
        Dim StdOut As System.IO.StreamReader = Nothing
        Dim StdErr As System.IO.StreamReader = Nothing

        If Diretorio = "" Then
            Diretorio = TemporaryDir()
        End If

        Dim Psi As New System.Diagnostics.ProcessStartInfo(Comando, Argumento)
        Psi.CreateNoWindow = True
        Psi.ErrorDialog = False
        Psi.UseShellExecute = False
        Psi.RedirectStandardError = True
        Psi.RedirectStandardInput = True
        Psi.RedirectStandardOutput = True
        Psi.WorkingDirectory = Diretorio

        If Usuario <> "" Then
            Psi.UserName = Usuario
        End If
        If Senha <> "" Then
            Psi.Password = New System.Security.SecureString
            For z As Integer = 1 To Len(Senha)
                Psi.Password.AppendChar(Mid(Senha, z, 1))
            Next
        End If
        If Dominio <> "" Then
            Psi.Domain = Dominio
        End If

        Proc = System.Diagnostics.Process.Start(Psi)
        StdIn = Proc.StandardInput
        StdOut = Proc.StandardOutput
        StdErr = Proc.StandardError
        If Entrada <> "" Then
            StdIn.WriteLine(Entrada)
        End If
        Result = StdOut.ReadToEnd()
        Erros = StdErr.ReadToEnd()
        Proc.WaitForExit(30)
        Return Result
    End Function

    Function FileExpr(ByVal ParamArray Segmentos() As String) As String
        Dim Arq As String = ExprExpr("\", "/", "", Segmentos)
        Return Arq
    End Function

    Sub CopyS(ByVal Origem As String, ByVal Destino As String)
        Dim Arquivo As String = System.IO.Path.GetFileName(Origem)
        Destino = FileExpr(Destino, Arquivo)
        System.IO.File.Copy(Origem, Destino)
        System.Console.WriteLine("Copiado: " & Origem & " para " & Destino)
    End Sub

    Function ExprExpr(ByVal Delim As String, ByVal DelimAlternativo As String, ByVal Inicial As Object, ByVal ParamArray Segmentos() As Object) As String
        Inicial = NZ(Inicial, "")
        Dim Lista As ArrayList = ParamArrayToArrayList(Segmentos)
        For Each item As Object In Lista
            If Not IsNothing(item) Then
                If Not IsNothing(DelimAlternativo) AndAlso DelimAlternativo <> "" Then
                    item = item.Replace(DelimAlternativo, Delim)
                End If
                item = NZ(item, "")
                If item <> "" Then
                    If Inicial <> "" Then
                        If Inicial.EndsWith(Delim) AndAlso item.StartsWith(Delim) Then
                            Inicial &= CType(item, String).Substring(Delim.Length)
                        ElseIf Inicial.EndsWith(Delim) OrElse item.StartsWith(Delim) Then
                            Inicial &= item
                        Else
                            Inicial &= Delim & item
                        End If
                    Else
                        Inicial &= item
                    End If
                End If
            End If
        Next
        Return Inicial
    End Function


    Function ParamArrayToArrayList(ByVal ParamArray Params() As Object) As Object

        ' caso não existam parâmetros
        If IsNothing(Params) OrElse Params.Length = 0 Then
            Return New ArrayList
        End If

        ' caso já seja um arraylist
        If Params.Length = 1 And TypeOf (Params(0)) Is ArrayList Then
            Return Params(0)
        End If

        ' caso tenha que juntar
        Dim ListaParametros As ArrayList = New ArrayList
        For Each Item As Object In Params
            If Not IsNothing(Item) Then

                ' >> TIPOS PREVISTOS EM ARRAYLIST...
                ' array
                ' arraylist
                ' string
                ' dataset
                ' datarowcollection

                If TypeOf Item Is Array Then
                    For Each SubItem As Object In Item
                        ListaParametros.AddRange(ParamArrayToArrayList(SubItem))
                    Next
                ElseIf TypeOf Item Is ArrayList OrElse Item.GetType.ToString.StartsWith("System.Collections.Generic.List") Then
                    ListaParametros.AddRange(Item)
                ElseIf TypeOf Item Is String Then
                    ListaParametros.Add(Item)
                ElseIf TypeOf Item Is DataSet Then
                    For Each Row As DataRow In Item.Tables(0).rows
                        For Each Campo As Object In Row.ItemArray
                            ListaParametros.Add(Campo)
                        Next
                    Next
                ElseIf TypeOf Item Is DataRow Then
                    For Each Campo As Object In CType(Item, DataRow).ItemArray
                        ListaParametros.Add(Campo)
                    Next
                ElseIf TypeOf Item Is System.IO.FileInfo Then
                    ListaParametros.Add(Item.name)
                Else
                    ListaParametros.Add(Item)
                End If
            End If
        Next
        Return ListaParametros
    End Function


    Function NZ(ByVal Valor As Object, Optional ByVal Def As Object = Nothing) As Object
        Dim tipo As String

        If Not IsNothing(Def) Then
            tipo = Def.GetType.ToString
        ElseIf IsNothing(Valor) Then
            Return Nothing
        Else
            tipo = Valor.GetType.ToString.Trim
        End If

        If IsNothing(Valor) OrElse IsDBNull(Valor) OrElse ((tipo = "System.DateTime" Or Valor.GetType.ToString = "System.DateTime") AndAlso Valor = CDate(Nothing)) Then
            Valor = Def
        End If

        Select Case tipo
            Case "System.Decimal"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Decimal)
                End If
                Return CType(Valor, Decimal)
            Case "System.String"
                If Valor.GetType.IsEnum Then
                    Return Valor.ToString
                End If
                Return CType(Valor, String)
            Case "System.Double"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Double)
                End If
                Return CType(Valor, Double)
            Case "System.Boolean"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return False
                End If
                Return CType(Valor, Boolean)
            Case "System.DateTime"
                Return CType(Valor, System.DateTime)
            Case "System.Single"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Single)
                End If
                Return CType(Valor, System.Single)
            Case "System.Byte"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Byte)
                End If
                Return CType(Valor, System.Byte)
            Case "System.Char"
                Return CType(Valor, System.Char)
            Case "System.SByte"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, SByte)
                End If
                Return CType(Valor, System.SByte)
            Case "System.Int32"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Int32)
                End If
                Return CType(Valor, Int32)
            Case "System.DBNull"
                Return Valor
            Case "System.Collections.ArrayList"
                Return ParamArrayToArrayList(Valor)
            Case "System.Data.DataSet"
                If IsNothing(Valor) Then
                    Return Def
                End If
                Return Valor
        End Select

        Return CType(Valor, String)
    End Function

    Function NZV(ByVal Valor As Object, Optional ByVal Def As Object = Nothing) As Object
        Dim Result As Object = NZ(Valor, Def)
        If TypeOf Result Is String AndAlso Result = "" Then
            Return Def
        ElseIf TypeOf Result Is Decimal AndAlso Result = 0 Then
            Return Def
        ElseIf TypeOf Result Is Double AndAlso Result = 0 Then
            Return Def
        ElseIf TypeOf Result Is Single AndAlso Result = 0 Then
            Return Def
        ElseIf TypeOf Result Is Int32 AndAlso Result = 0 Then
            Return Def
        ElseIf TypeOf Result Is Byte AndAlso Result = 0 Then
            Return Def
        End If
        Return Result
    End Function

    Class Arq
        Public Arquivo As String
        Public Cria As Date
        Sub New(ByVal Arquivo As String)
            Me.Arquivo = Arquivo
        End Sub
    End Class

    Public Function ListagemArq(ByVal Diretorio As String, Optional ByVal Criterio As String = "*.*", Optional ByVal SubDir As Boolean = False, Optional ByVal Ordenar As Boolean = False) As List(Of Arq)
        Dim ret As New List(Of Arq)
        If Diretorio <> "" Then
            For Each fl As String In System.IO.Directory.GetFiles(Diretorio, Criterio)
                ret.Add(New Arq(fl))
            Next
            If SubDir Then
                For Each dr As String In System.IO.Directory.GetDirectories(Diretorio)
                    ret.AddRange(ListagemArq(dr, Criterio))
                Next
            End If
        End If
        For Each A As Arq In ret
            Try
                A.Cria = System.IO.File.GetCreationTime(A.Arquivo)
            Catch
                A.Cria = Nothing
            End Try
        Next
        If Ordenar Then
            ret = (From x In ret Order By x.Cria Descending, x.Arquivo Descending).ToList
        End If
        Return ret
    End Function


    Public Function VersaoApl() As String
        Return "V" & Format(My.Application.Info.Version.Major, "00") & "." & Format(My.Application.Info.Version.Minor, "00") & "." & Format(My.Application.Info.Version.MajorRevision, "00") & "." & Format(My.Application.Info.Version.MinorRevision, "00")
    End Function

    Function SemAspas(ByVal Texto As String) As String
        Return Texto.Trim("""", Chr(147), Chr(148))
    End Function

    Dim Help As Boolean = False
    Dim NoDir As Integer = 0
    Dim Semana As Integer = 0
    Dim Mes As Integer = 0
    Dim Ano As Integer = 0
    Dim Caminho As String = ""
    Dim Comando As String = ""

    Sub Main()
        Try
            Dim Args As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs()

            For z As Integer = 0 To Args.Count - 1 Step 2
                Dim Coma As String = SemAspas(Args(z)).ToLower

                Select Case Coma
                    Case "-help"
                        Help = True
                End Select
            Next

            If Help Or Args.Count = 0 Then
                Dim Msg As String = vbCrLf & vbCrLf & "RetemApenas " & VersaoApl() & " - " & "Deixa apenas quantidade de arquivos especificada" & vbCrLf & vbCrLf
                Msg &= "     Help:               -help" & vbCrLf & vbCrLf
                Msg &= "     RetemApenas 6 4 3 2 c:\diret\arqx* [cmd]     onde..." & vbCrLf & vbCrLf
                Msg &= "     6     sign o numero de arquivos que serao mantidos no diretorio" & vbCrLf
                Msg &= "     4     sign o numero de arquivos mantidos semanalmente" & vbCrLf
                Msg &= "     3     sign o numero de arquivos mantidos mensalmente" & vbCrLf
                Msg &= "     2     sign o numero de arquivos mantidos anualmente" & vbCrLf
                Msg &= "     c:\diret\arq*   sign local e mascara a considerar" & vbCrLf & vbCrLf
                Msg &= "     [cmd] corresponde ao comando a ser executado apos o tratamento" & vbCrLf & vbCrLf
                Msg &= "O programa criara estrutura semanal, mensal e anual mantendo" & vbCrLf
                Msg &= "    a quantidade de arquivos indicada em cada nivel." & vbCrLf
                Msg &= "    Indicar 0 (ZERO) quando nao desejar aquele nivel." & vbCrLf & vbCrLf
                Msg &= "Exemplo:" & vbCrLf
                Msg &= "    retemapenas 4 3 2 1 C:\_LUC\BACK\T*.*"
                Msg &= vbCrLf
                System.Console.WriteLine(Msg)
                Exit Sub
            End If

            NoDir = SemAspas(Args(0))
            If Args.Count > 1 Then
                Semana = SemAspas(Args(1))
                If Args.Count > 2 Then
                    Mes = SemAspas(Args(2))
                    If Args.Count > 3 Then
                        Ano = SemAspas(Args(3))
                        If Args.Count > 4 Then
                            Caminho = SemAspas(Args(4))
                            If Args.Count > 5 Then
                                Comando = SemAspas(Args(5))
                            End If
                        End If
                    End If
                End If
            End If

            If NoDir = 0 And Semana = 0 And Mes = 0 And Ano = 0 Then
                Throw New Exception("Pelo menos um arquivo devera ser mantido.")
            End If
            If Caminho = "" Then
                Throw New Exception("Informe caminho.")
            End If

            Dim Diretorio As String = System.IO.Path.GetDirectoryName(Caminho)
            Dim Mascara As String = System.IO.Path.GetFileName(Caminho)
            Dim Lista As List(Of Arq) = ListagemArq(Diretorio, Mascara, , True)

            If Lista.Count = 0 Then
                Throw New Exception("Nao existem arquivos que coincidam com a mascara no caminho especificado.")
            End If


            Dim DirAno As String = FileExpr(Diretorio, "Anual")
            Dim DirMes As String = FileExpr(Diretorio, "Mensal")
            Dim DirSemana As String = FileExpr(Diretorio, "Semanal")

            If Not System.IO.Directory.Exists(DirAno) Then
                MkDir(DirAno)
            End If
            If Not System.IO.Directory.Exists(DirMes) Then
                MkDir(DirMes)
            End If
            If Not System.IO.Directory.Exists(DirSemana) Then
                MkDir(DirSemana)
            End If

            If Semana > 0 Then
                Dim ListaSem As List(Of Arq) = ListagemArq(DirSemana, Mascara, , True)
                If ListaSem.Count = 0 OrElse DateDiff(DateInterval.Day, ListaSem(0).Cria, Lista(0).Cria) >= 7 Then
                    CopyS(Lista(0).Arquivo, DirSemana)
                End If
                ListaSem = ListagemArq(DirSemana, Mascara, , True)
                For z As Integer = Semana To ListaSem.Count - 1
                    System.IO.File.Delete(ListaSem(z).Arquivo)
                    System.Console.WriteLine("Eliminado: " & ListaSem(z).Arquivo)
                Next
            End If

            If Mes > 0 Then
                Dim ListaMes As List(Of Arq) = ListagemArq(DirMes, Mascara, , True)
                If ListaMes.Count = 0 OrElse DateDiff(DateInterval.Month, ListaMes(0).Cria, Lista(0).Cria) >= 1 Then
                    CopyS(Lista(0).Arquivo, DirMes)
                End If
                ListaMes = ListagemArq(DirMes, Mascara, , True)
                For z As Integer = Mes To ListaMes.Count - 1
                    System.IO.File.Delete(ListaMes(z).Arquivo)
                    System.Console.WriteLine("Eliminado: " & ListaMes(z).Arquivo)
                Next
            End If

            If Ano > 0 Then
                Dim ListaAno As List(Of Arq) = ListagemArq(DirAno, Mascara, , True)
                If ListaAno.Count = 0 OrElse DateDiff(DateInterval.Month, ListaAno(0).Cria, Lista(0).Cria) >= 1 Then
                    CopyS(Lista(0).Arquivo, DirAno)
                End If
                ListaAno = ListagemArq(DirAno, Mascara, , True)
                For z As Integer = Ano To ListaAno.Count - 1
                    System.IO.File.Delete(ListaAno(z).Arquivo)
                    System.Console.WriteLine("Eliminado: " & ListaAno(z).Arquivo)
                Next
            End If

            For z As Integer = NoDir To Lista.Count - 1
                System.IO.File.Delete(Lista(z).Arquivo)
                System.Console.WriteLine("Eliminado: " & Lista(z).Arquivo)
            Next

            If Comando <> "" Then
                System.Console.Write("Exec comando:" & vbCrLf)
                System.Console.Write(Comando)
                Dim Erros As String = ""
                Dim Ret As String = DosShell("cmd", "/c " & Replace(Comando, "'", """"), "", "", Erros)
                System.Console.WriteLine(Ret)
                If Erros <> "" Then
                    Throw New Exception(Erros & " ao executar comando " & Comando)
                End If
            End If

        Catch EX As Exception
            System.Console.WriteLine("[ERRO] " & EX.Message)
        End Try

    End Sub

End Module

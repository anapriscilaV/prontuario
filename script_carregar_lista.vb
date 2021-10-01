
Private Sub btn_pesquisa_Click()
        
Dim strSQL As String
Dim i As Long
   
'Consulta cadastro animal dentro da planilha com os dados de origem

    With Planilha5.Range("C:C")
 
        'Aqui onde será digitado o que procurar
 
        Set c = .Find(campo_pesquisa.Value, LookIn:=xlValues, LOOKAT:=xlWhole)
                         
 
        'Aqui outra tomada de decisão, caso encontre sua pesquisa, será retornado nas caixas de textbox.
        If Not c Is Nothing Then
        
        
            TextBox38.Text = c.Offset(0, 3) 'carregar nome animal
            TextBox29.Text = c.Offset(0, 4) 'carregar GRUPO/ ESPÉCIE:
            TextBox30.Text = c.Offset(0, 15) 'carregar IDENTIFICAÇÃO:
            TextBox40.Text = c.Offset(0, 14) 'carregar Tipo Marcação
            'TextBox31.Text = C.Offset(0, 11) 'carregar DATA DE RESGATE:
            TextBox32.Text = c.Offset(0, 11) 'carregar DATA DE NASCIMENTO:
            TextBox28.Text = c.Offset(0, 16) 'carregar PELAGEM:
            TextBox39.Text = c.Offset(0, 8) 'carregar SEXO ANIMAL
            TextBox35.Text = c.Offset(0, 2) 'Nome Tutor
            CheckBox1_zas.Value = False
            CheckBox3_zss.Value = True
            CheckBox2_excepcional = False
            
            
'Consultando e populando o listbox de atendimento
'implementar lógica para query MYSQL
'INSERIR LOOP MultD

'strSQL = "SELECT av_data from dt_atendimento WHERE animal_code LIKE '%"&campo_pesquisa&"%'"

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Set cn = New ADODB.Connection
    cn.ConnectionString = "DRIVER={MySQL ODBC 8.0 ANSI Driver};" _
                              & "SERVER= 187.45.196.151;" _
                              & "DATABASE=valefauna;" _
                              & "UID=valefauna;" _
                              & "PASSWORD=a88KSfH3hyR@xe;Option=3;"

cn.Open
Set rs = New ADODB.Recordset
rs.ActiveConnection = cn

rs.Source = "Select * from dt_atendimento where animal_code like '%" & c.Value & "%'"

rs.Open
    'filtro_data.List = rs.GetRows
    'loop para filtrar data
    'casar caixa de combinação
    
    With listagem_atendimento
    .ColumnCount = rs.Fields.Count
        listagem_atendimento.List = WorksheetFunction.Transpose(rs.GetRows)
        
'listagem_atendimento.AddItem "Abrindo Consulta dt_atendimento"
'WorkesheetFunction.Transpose(rs.GetRows)
'rs.Close

'rs.Close
        
        
        Dim dt As ADODB.Recordset

Set dt = New ADODB.Recordset
dt.ActiveConnection = cn
dt.Source = "SELECT av_data from dt_atendimento WHERE animal_code LIKE '%" & c.Value & "%'"
dt.Open
             
    With filtro_data
    .ColumnCount = dt.Fields.Count
        filtro_data.List = WorksheetFunction.Transpose(dt.GetRows)
        'listagem_atendimento.SelectedIndex = listagem_atendimento.Items.IndexOf(filtro_data.SelectedItem)
        
'listagem_atendimento.AddItem "Abrindo Consulta dt_atendimento"
'WorkesheetFunction.Transpose(rs.GetRows)
'rs.Close
cn.Close
'rs.Close
        End With
        
        
End With
End If
        End With
End Sub
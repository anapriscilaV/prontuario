
# coding: utf8

from os import error
from typing import final
import mysql.connector
from mysql.connector import Error


#inserção dados excel no BD

try:
  con = mysql.connector.connect(host='localhost',database='dbo.prontuario',
                                user='root',password='')
  inserir_dados = """INSERT INTO `dbo.prontuario`.`dt_prontuario`
  (id_veterinario, animal_code, id_tipo_atendimento, id_local, zona_id, motivo_atendimento, exame_clinico_anotacoes, exames_complementares,
    tratamento_protocolo, prognostico, observacao_prontuario) 
    
    for dados_formulario in range (1, sheet.nrows):
        header_list_animal = ['GRUPO/ ESPÉCIE:', 'NOME ANIMAL:','IDENTIFICAÇÃO: ', 'DATA DE RESGATE:','DATA DE NASCIMENTO: ','PELAGEM:']
        header_list_tutor = ['NOME TUTOR', 'CÓDIGO:', 'ENDEREÇO:']
        header_list_atendimento =['DATA DO ATENDIMENTO:', 'VETERINÁRIO RESPONSÁVEL/ CRMV INSTALAÇÃO:', 'PESO:',  'ECC:', 'QUEIXA PRINCIPAL/ MOTIVO DO ATENDIMENTO:', 'EXAME CLÍNICO:', 'SUSPEITA CLÍNICA/ DIAGNÓSTICO:', 'TRATAMENTO/ PROTOCOLO:' , 'OBSERVAÇÕES:']
        
    ##comparação entre COLUNA BD COM NOME DO CABEÇALHO DEFINIDO 
    
    
    VALUES 
    (20, 75,5, 1,'TESTE Inserção Script Python' ,'TESTE Inserção Script Python','TESTE Inserção Script Python','TESTE Inserção Script Python',
    'TESTE Inserção Script quebrar primeiro nome do veterinario', 'TESTE Inserção Script Python', 'TESTE Inserção Script Python');
  """
  
  cursor=con.cursor()
  cursor.execute(inserir_dados)
  con.commit()
  print(cursor.rowcount,"registros inseridos na tabela")
  
  print('Dados inseridos no BD!')
  
except Error as erro:
  print('Falha ao inserir dados: {}'.format(erro))
finally:
  if(con.is_connected()):
        con.close()
        cursor.close
        print('Conexão Encerrada')
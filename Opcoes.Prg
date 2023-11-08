#include <hmg.ch>

Procedure config()

          if IsWIndowActive(Opcoes)
             DoMethod( "Opcoes","SETFOCUS" )
          else
             LOAD WINDOW Opcoes
             CENTER WINDOW Opcoes
             ACTIVATE WINDOW Opcoes
          end

Return

Procedure opcoes_salvar()
          Local i, nFrequencia

          if ( AllTrim(Opcoes.Text_senha.Value) == g_aUsuarios[Opcoes.Combo_usuarios.Value]['senha'] )

             RegistryWrite( g_cRegPath + "Monitoring\das", Opcoes.Text_das.Value )
             RegistryWrite( g_cRegPath + "Monitoring\ate", Opcoes.Text_as.Value )
             RegistryWrite( g_cRegPath + "Monitoring\frequencia", Opcoes.Text_frequencia.Value )

             Opcoes.Release

          else
             MsgExclamation('Senha Inválida, favor conferir!','Senha')
          end

Return

Procedure opcoes_cancelar()
          Opcoes.Release
Return

Function ver_senha_action()

         Opcoes.Label_ver_senha.Value := Opcoes.Text_senha.Value
         Opcoes.Label_ver_senha.Visible := .T.
         inkey(1.5)
         Opcoes.Label_ver_senha.Visible := .F.

Return Nil

Procedure opcoes_form_oninit()
          Local i

          SetProperty( "Main", "Timer_Mail", "Enabled", .F. )
          SetProperty( "Main", "Timer_yes_update", "Enabled", .F. )

          if !( g_oEmpresas:QtdeEmpresas() == 0 )
             Opcoes.Button_opcoes_salvar.Enabled := .T.
             Opcoes.Text_das.Value   := RegistryRead( g_cRegPath + 'Monitoring\das' )
             Opcoes.Text_as.Value    := RegistryRead( g_cRegPath + 'Monitoring\ate' )
             Opcoes.Text_frequencia.Value := RegistryRead( g_cRegPath + 'Monitoring\frequencia' )
             MsgStatus( LTrim(Str(g_oEmpresas:QtdeEmpresas())) + if(g_oEmpresas:QtdeEmpresas() > 1, ' registros carregados.', ' registro carregado.') )
          else
             Opcoes.Button_opcoes_salvar.Enabled := .F.
             MsgStatus('Nenhum registro encontrado')
             MsgExclamation('Nenhum registro encontrado', 'eMailCTe')
          end

          if HMG_LEN(g_aUsuarios) > 0
             Opcoes.Combo_usuarios.Enabled := .F.
             Opcoes.Combo_usuarios.DeleteAllItems
             FOR i := 1 TO HMG_LEN(g_aUsuarios)
                 Opcoes.Combo_usuarios.AddItem(g_aUsuarios[i]['login'])
             NEXT
             Opcoes.Combo_usuarios.Value   := 1
             Opcoes.Combo_usuarios.Enabled := .T.
          end

          MsgStatus()

Return

Procedure opcoes_Text_frequencia_onlostfocus()
          if Opcoes.Text_frequencia.Value < 1
             Opcoes.Text_frequencia.Value := 1
          end
Return

Procedure opcoes_text_das_onlostfocus()
          if Opcoes.Text_das.Value < "00:00" .or. Opcoes.Text_das.Value > "23:59"
             Opcoes.Text_das.Value := "  :  "
          end
Return
Procedure opcoes_text_as_onlostfocus()
          if Opcoes.Text_as.Value < "00:00" .or. Opcoes.Text_as.Value > "23:59"
             Opcoes.Text_as.Value := "  :  "
          end
Return

Procedure opcoes_form_onrelease()
          SetProperty( "Main", "Timer_Mail", "Enabled", .T. )
          SetProperty( "Main", "Timer_yes_update", "Enabled", .T. )
Return

#include <hmg.ch>

Procedure MonitoraErros

          if IsWIndowActive(Opcoes)
             Return
          end

          // Verifica se algum sistema da Sistrom registrou erros e envia por e-mail o log de erros para suporte Sistrom

          if ( Seconds() - g_iTimer_Erro ) >= 7200  // segundos = 2 Horas
             CheckErrorsLog()
             g_iTimer_Erro := Seconds()
          end

Return

Procedure CheckErrorsLog
          Local aErrorLog := {}
          Local sCliente

          if !hb_DirExists( '..\Log_Erros\' )
             DirMake( '..\Log_Erros\' )
          end

          if hb_FileExists( 'ErrorLog.Htm' )
             if hb_FileExists( '..\Log_Erros\ErrorLog_emcte.Htm' )
                fErase( '..\Log_Erros\ErrorLog_emcte.Htm' )
             end
             fRename( 'ErrorLog.Htm', '..\Log_Erros\ErrorLog_emcte.Htm' )
             sCliente := RegistryRead( 'HKEY_CURRENT_USER\SOFTWARE\Sistrom\eMailCTE\Cliente' )
             AADD( aErrorLog, {'Sistema' => "eMailCTE", 'Cliente' => sCliente, 'Arquivo' => "..\Log_Erros\ErrorLog_emcte.Htm"} )
          end

          if hb_FileExists( '..\PrintCTE\ErrorLog.Htm' )
             if hb_FileExists( '..\Log_Erros\ErrorLog_ptcte.Htm' )
                fErase( '..\Log_Erros\ErrorLog_ptcte.Htm' )
             end
             fRename( '..\PrintCTE\ErrorLog.Htm', '..\Log_Erros\ErrorLog_ptcte.Htm' )
             sCliente := RegistryRead( 'HKEY_CURRENT_USER\SOFTWARE\Sistrom\PrintCTE\Cliente' )
             AADD( aErrorLog, {'Sistema' => "PrintCTE", 'Cliente' => sCliente, 'Arquivo' => "..\Log_Erros\ErrorLog_ptcte.Htm"} )
          end

          if hb_FileExists( '..\PrintCTE_LWCargas\ErrorLog.Htm' )
             if hb_FileExists( '..\Log_Erros\ErrorLog_ptlwcte.Htm' )
                fErase( '..\Log_Erros\ErrorLog_ptlwcte.Htm' )
             end
             fRename( '..\PrintCTE_LWCargas\ErrorLog.Htm', '..\Log_Erros\ErrorLog_ptlwcte.Htm' )
             sCliente := RegistryRead( 'HKEY_CURRENT_USER\SOFTWARE\Sistrom\PrintCTE_LWCargas\Cliente' )
             if !( sCliente == NIL )
                AADD( aErrorLog, {'Sistema' => "PrintCTE_LWCargas", 'Cliente' => sCliente, 'Arquivo' => "..\PrintCTE_LWCargas\ErrorLog_ptlwcte.Htm"} )
             end
          end

          if hb_FileExists( '..\Atualizacao_remota\ErrorLog.Htm' )
             if hb_FileExists( '..\Log_Erros\ErrorLog_rupdt.Htm' )
                fErase( '..\Log_Erros\ErrorLog_rupdt.Htm' )
             end
             fRename( '..\Atualizacao_remota\ErrorLog.Htm', '..\Log_Erros\ErrorLog_rupdt.Htm' )
             sCliente := RegistryRead( 'HKEY_CURRENT_USER\SOFTWARE\Sistrom\RemoteUpdate\Cliente' )
             AADD( aErrorLog, {'Sistema' => "RemoteUpdate", 'Cliente' => sCliente, 'Arquivo' => "..\Atualizacao_remota\ErrorLog_rupdt.Htm"} )
          end

          if hb_FileExists( '..\TMS_Cloud_CTE\ErrorLog.Htm' )
             if hb_FileExists( '..\Log_Erros\ErrorLog_tmscte.Htm' )
                fErase( '..\Log_Erros\ErrorLog_tmscte.Htm' )
             end
             fRename( '..\TMS_Cloud_CTE\ErrorLog.Htm', '..\Log_Erros\ErrorLog_tmscte.Htm' )
             sCliente := RegistryRead( 'HKEY_CURRENT_USER\SOFTWARE\Sistrom\TMS.Cloud\Cliente' )
             AADD( aErrorLog, {'Sistema' => "TMSCloudCTE", 'Cliente' => sCliente, 'Arquivo' => "..\TMS_Cloud_CTE\ErrorLog_tmscte.Htm"} )
          end

          if !( HMG_LEN( aErrorLog ) == 0 )
             EnviaErros( aErrorLog )
          end

Return

Procedure EnviaErros( aErrors )
          Local i
          Local aZips := {}, hErro
          Local email, msg
          g_oEmpresas:SetEmpresa(1)

          email := Tsmtp_email():new(g_oEmpresas:smtp_servidor, g_oEmpresas:smtp_porta)
          email:setLogin(g_oEmpresas:smtp_email, g_oEmpresas:smtp_login, g_oEmpresas:smtp_senha, g_oEmpresas:smtp_pass)
          email:setRecipients('suporte@sistrom.com.br')
          msg := 'Erro(s) ocorrido(s) no(s) sistema(s) Desktop da Sistrom.' + CRLF
          msg += 'Coletado em: ' + DTOC(Date()) + ' as ' + Time() + CRLF + CRLF

          hErro := aErrors[1]

          FOR EACH hErro IN aErrors
              AADD( aZips, hErro['Arquivo'] )
              msg += HMG_PADR(hErro['Sistema'], 20 ) + HMG_PADR( hErro['Cliente'], 30 ) + hErro['Arquivo'] + CRLF
          NEXT EACH

          if hb_FileExists( '..\Log_Erros\ErrorLog.zip' )
             fErase( '..\Log_Erros\ErrorLog.zip' )
          end

          if !Empty(aZips)

				 CompressFiles( '..\Log_Erros\ErrorLog.zip', aZips )

             if hb_FileExists( '..\Log_Erros\ErrorLog.zip' )
                MsgStatus('Enviando arquvivos de Error Logs', 'emailSend' )
                RegistraLog( 'Enviando arquvivos de Error Logs' )
                email:setMsg('Errorlog File - ' + hb_HGetDef(hErro, 'Cliente', ''), msg)
                AADD( email:attachment, '..\Log_Erros\ErrorLog.zip' )
                email:sendmail()
             end

			 end

Return
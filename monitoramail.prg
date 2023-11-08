#include <hmg.ch>

DECLARE WINDOW Main
DECLARE WINDOW Opcoes

Procedure MonitoraMails()
          Local nSegundos := RegistryRead( g_cRegPath + 'Monitoring\frequencia' ) * 60

          if IsWIndowActive(Opcoes)
             Return
          end

          if HB_ULEFT ( Time(), 2 ) == "23" .or. HB_ULEFT ( Time(), 2 ) == "00"
				 RegistraLog('eMailCTE ' + g_cVersao + ': Desligamento do sitema programado às 23hs, período de inatividade até 01hs, virada do dia.' )
				 RELEASE WINDOW ALL
			 end

			 if ( Seconds() - g_iTimer ) >= nSegundos
             GetMailsWeb()
             g_iTimer := Seconds()
          end

Return

Procedure GetMailsWeb()
          Local sql
          Local i, n := 1, nStatus
          Local oQry

          if ( g_oEmpresas:QtdeEmpresas() == 0 )
             RegistraLog('Qtde de Empresas é 0, retornou sem monitorar e-mails' )
             Return
          end

          MsgStatus( 'Monitorando eMails web...' )

          sql := "SELECT cte_id, "
          sql += "emp_id, "
          sql += "clie_remetente_id, "
          sql += "clie_coleta_id, "
          sql += "clie_expedidor_id, "
          sql += "clie_recebedor_id, "
          sql += "clie_destinatario_id, "
          sql += "clie_entrega_id, "
          sql += "clie_tomador_id, "
          sql += "cte_numero, "
          sql += "cte_chave, "
          sql += "cte_situacao, "
          sql += "cte_pdf, "
          sql += "cte_cancelado_pdf, "
          sql += "cte_xml, "
          sql += "cte_cancelado_xml, "
          sql += "cte_exibe_consulta_cliente AS enviar_email_cliente "
          sql += "FROM ctes "
          sql += "WHERE emp_id IN ("

          FOR i := 1 TO g_oEmpresas:QtdeEmpresas()

              g_oEmpresas:SetEmpresa(i)

              if n > 1
                 sql += ","
              end
              sql += LTrim(STR(g_oEmpresas:id))
              n++

          NEXT i

          if (n == 1)
             // Não tem empresas para monitoramento
             RegistraLog('n == 1, Não tem empresas para monitoramento, retornou sem monitorar e-mails' )
             MsgStatus()
             Return
          end

          sql += ") AND cte_chave IS NOT NULL AND cte_chave != '' AND cte_email_enviado = 0 "
          sql += "AND ((cte_situacao = 'AUTORIZADO' AND cte_pdf IS NOT NULL AND cte_pdf != '') OR (cte_situacao = 'CANCELADO' AND cte_cancelado_pdf IS NOT NULL AND cte_cancelado_pdf != '')) "
          sql += "ORDER BY emp_id, cte_numero;"

          oQry := ExecutaQuery(sql)

          if ExecutouQuery( @oQry, sql )

             if oQry:NetErr()

                RegistraLog('Sem conexão com a internet. Erro SQL: ' + oQry:Error())
                PlayExclamation()
                MsgExclamation('Sem conexão com a internet. Erro SQL: ' + CRLF + oQry:Error(), 'eMailCTe: Carregando CTes')
                oQry:Destroy()
                RELEASE WINDOW ALL

             else

                if !(oQry:LastRec() == 0)
                  send_emails(oQry)
                  upload_mail_events()
                  upload_mail_errors()
                  MsgStatus()
                end

             end

             oQry:Destroy()

          else
             RegistraLog('Solicitação SQL ao Servidor foi perdida' + CRLF + '| SQL: ' + sql,, true)
             MsgStatus( 'Solicitação SQL ao Servidor foi perdida', 'dbOff' )
             PlayExclamation()
             MsgExclamation('Solicitação SQL ao Servidor foi perdida!' + CRLF + "Tente mais tarde, se persistir, chame o suporte.", 'eMailCTe: Monitorando eMail(s)')
             RELEASE WINDOW ALL
          end

Return

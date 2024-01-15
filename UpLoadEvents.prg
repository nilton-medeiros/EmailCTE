#include <hmg.ch>

Procedure upload_mail_events()
          Local i := 0, nId_cte := 0, nEMail := 1
          Local oQuery
          Local hMail
          Local mail_sql := "UPDATE ctes SET cte_email_enviado=1 WHERE cte_id IN ("
          Local evts_sql := "INSERT INTO ctes_eventos (cte_id, cte_ev_protocolo, cte_ev_data_hora, cte_ev_evento, cte_ev_detalhe) VALUES "

          MsgStatus('Atualizando eventos dos e-Mails na web...', 'dbRefresh')

          FOR EACH hMail IN g_aMaiLogEvent

              i++

              if !( nId_cte == hMail['cte_id'] )

                 nId_cte := hMail['cte_id']

                 if !(nEMail == 1)
                    mail_sql += ","
                 end

                 mail_sql += hb_ntos(hMail['cte_id'])
                 nEMail++

              end

              evts_sql += if( (i==1), [(], [, (] )                                        // Inicio dos VALUES
              evts_sql += LTrim(STR(hMail['cte_id'])) + ", "                              // cte_id
              evts_sql += "'eMailCTE', "                                                  // cte_ev_protocolo
              evts_sql += "'" + DateToSQL(hMail['data']) + " " + hMail['hora'] + "', "    // cte_ev_data_hora
              evts_sql += "'EMAI', "                                                       // cte_ev_evento
              evts_sql += "'ENVIO DE e-MAILs: " + mysql_escape( hMail['mensagem'] ) + "')"    // cte_ev_detalhe e fechamento dos VALUES

          NEXT EACH

          // Da baixa nos ctes que foram enviado emails

          if !(i == 0)

             // Houve eMails enviados, dar baixa no BD

             mail_sql += ") AND cte_email_enviado=0"

             oQuery := ExecutaQuery(mail_sql)

             if ExecutouQuery(@oQuery, mail_sql)
                if oQuery:NetErr()
                   RegistraLog('Erro SQL: ' + oQuery:Error() + ' | Descrição do comando SQL: ' + CRLF + '| SQL: ' + mail_sql)
                   MsgStatus('Erro SQL', 'dbError')
                   PlayExclamation()
                   MsgExclamation('Descrição do erro: ' + CRLF + oQuery:Error(), 'eMailCTe: Erro SQL')
                   MsgExclamation({'Comando: ', CRLF, mail_sql, CRLF, 'Ligue para o Suporte ou reiniciar o programa'}, 'eMailCTe: Comando SQL')
                   oQuery:Destroy()
                   RELEASE WINDOW ALL
                end
                oQuery:Destroy()
             else
                RegistraLog('Solicitação SQL ao Servidor foi perdida | ' + "COMANDO SQL: " + CRLF + '| SQL: ' + mail_sql)
                MsgStatus('Solicitação SQL ao Servidor foi perdida', 'dbError')
                PlayExclamation()
                MsgStop('Solicitação SQL ao Servidor foi perdida' + CRLF + "COMANDO SQL: " + CRLF + mail_sql, 'eMailCTe: Atualizando Base de dados')
                RELEASE WINDOW ALL
             end
             RegistraLog('UPDATE SQL: ' + mail_sql )
          else
            RegistraLog('Array g_aMaiLogEvent retornou vazio!')
          end

          if !(i == 0)
            // Lança os eventos relacionados a baixa dos PDFs & MXLs

            oQuery := ExecutaQuery(evts_sql)

            if ExecutouQuery(@oQuery, evts_sql)
               if oQuery:NetErr()
                  RegistraLog('Erro SQL: ' + oQuery:Error() + ' | Descrição do comando SQL: ' + CRLF + '| SQL: ' + evts_sql)
                  MsgStatus('Erro SQL', 'dbError')
                  PlayExclamation()
                  MsgExclamation('Descrição do erro: ' + CRLF + oQuery:Error(), 'eMailCTe: Erro SQL')
                  MsgExclamation({'Comando: ', CRLF, evts_sql, CRLF, 'Ligue para o Suporte ou reiniciar o programa'}, 'eMailCTe: Comando SQL')
                  oQuery:Destroy()
                  RELEASE WINDOW ALL
               end
               oQuery:Destroy()

            else
               RegistraLog('Solicitação SQL ao Servidor foi perdida | COMANDO SQL: ' + CRLF + '| SQL: ' + evts_sql)
               MsgStatus('Solicitação SQL ao Servidor foi perdida', 'dbError')
               PlayExclamation()
               MsgStop('Solicitação SQL ao Servidor foi perdida' + CRLF + "COMANDO SQL: " + CRLF + evts_sql, 'eMailCTe: Atualizando Base de dados')
               RELEASE WINDOW ALL
            end
            RegistraLog('INSERT SQL: ' + evts_sql )
          else
            RegistraLog('Array g_aMaiLogEvent retornou vazio!')
          endif

          g_aMaiLogEvent := {}     //  Limpa e reinicia o log de enventos do eMail

          MsgStatus()

Return

Procedure upload_mail_errors()
          Local i := 0, nId_cte := 0, nEMail := 1
          Local oQuery, hMail
          Local mail_sql := "UPDATE ctes SET cte_arquivos_baixados=0 WHERE cte_id IN ("
          Local evts_sql := "INSERT INTO ctes_eventos (cte_id, cte_ev_protocolo, cte_ev_data_hora, cte_ev_evento, cte_ev_detalhe) VALUES "

          MsgStatus('Atualizando eventos dos e-Mails na web...', 'dbRefresh')

          FOR EACH hMail IN g_aMaiComErros

              i++

              if !( nId_cte == hMail['cte_id'] )

                 nId_cte := hMail['cte_id']

                 if !(nEMail == 1)
                    mail_sql += ","
                 end

                 mail_sql += hb_ntos(hMail['cte_id'])
                 nEMail++

              end

              evts_sql += if( (i==1), [(], [, (] )                                     // Inicio dos VALUES
              evts_sql += LTrim(STR(hMail['cte_id'])) + ", "                           // cte_id
              evts_sql += "'eMailCTE', "                                               // cte_ev_protocolo
              evts_sql += "'" + DateToSQL(hMail['data']) + " " + hMail['hora'] + "', " // cte_ev_data_hora
              evts_sql += "'EMAI', "                                                    // cte_ev_evento
              evts_sql += "'SENT e-MAILs: " + mysql_escape( hMail['mensagem'] ) + "')" // cte_ev_detalhe e fechamento dos VALUES

          NEXT EACH

          // Da baixa nos ctes que foram enviado emails

          if (nEMail > 1)

             // Houve eMails enviados, dar baixa no BD

             mail_sql += ")"

             oQuery := ExecutaQuery(mail_sql)

             if ExecutouQuery(@oQuery, mail_sql)
                if oQuery:NetErr()
                   RegistraLog('Erro SQL: ' + oQuery:Error() + ' | Descrição do comando SQL: ' + CRLF + '| SQL: ' + mail_sql)
                   MsgStatus('Erro SQL', 'dbError')
                   PlayExclamation()
                   MsgExclamation('Descrição do erro: ' + CRLF + oQuery:Error(), 'eMailCTe: Erro SQL')
                   MsgExclamation({'Comando: ', CRLF, mail_sql, CRLF, 'Ligue para o Suporte ou reiniciar o programa'}, 'eMailCTe: Comando SQL')
                   oQuery:Destroy()
                   RELEASE WINDOW ALL
                end
                oQuery:Destroy()
             else
                RegistraLog('Solicitação SQL ao Servidor foi perdida | ' + "COMANDO SQL: " + CRLF + '| SQL: ' + mail_sql)
                MsgStatus('Solicitação SQL ao Servidor foi perdida', 'dbError')
                PlayExclamation()
                MsgStop('Solicitação SQL ao Servidor foi perdida', 'eMailCTe: Atualizando Base de dados')
                RELEASE WINDOW ALL
             end
             RegistraLog('UPDATE SQL: ' + mail_sql)
//        else
//          RegistraLog('Array g_aMaiComErros retornou vazio!')
          end

          if !(i == 0)

             // Lança os eventos relacionados a baixa dos PDFs & MXLs

             oQuery := ExecutaQuery(evts_sql)

             if ExecutouQuery(@oQuery, evts_sql)
                if oQuery:NetErr()
                   RegistraLog('Erro SQL: ' + oQuery:Error() + ' | Descrição do comando SQL: ' + CRLF + '| SQL: ' + evts_sql)
                   MsgStatus('Erro SQL', 'dbError')
                   PlayExclamation()
                   MsgExclamation('Descrição do erro: ' + CRLF + oQuery:Error(), 'eMailCTe: Erro SQL')
                   MsgExclamation({'Comando: ', CRLF, evts_sql, CRLF, 'Ligue para o Suporte ou reiniciar o programa'}, 'eMailCTe: Comando SQL')
                   oQuery:Destroy()
                   RELEASE WINDOW ALL
                end
                oQuery:Destroy()
                RegistraLog('INSERT SQL: ' + evts_sql)

             else
                RegistraLog('Solicitação SQL ao Servidor foi perdida | COMANDO SQL: ' + CRLF + '| SQL: ' + evts_sql)
                MsgStatus('Solicitação SQL ao Servidor foi perdida', 'dbError')
                PlayExclamation()
                MsgStop('Solicitação SQL ao Servidor foi perdida', 'eMailCTe: Atualizando Base de dados')
                RELEASE WINDOW ALL
             end
             RegistraLog('INSERT SQL: ' + evts_sql)
//       else
//          RegistraLog('Array g_aMaiComErros retornou vazio!')
         endif
          g_aMaiComErros := {}     //  Limpa e reinicia o log de enventos do eMail

          MsgStatus()

Return

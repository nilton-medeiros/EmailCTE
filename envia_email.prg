#include <hmg.ch>

#define true  .T.
#define false .F.

// Atualizado 2022-05-31 12:30
Procedure envia_emails( oQuery )
          Local cEmitente, cFone, eMailComercial, eMailContabil, cPortal, cMail, cIPEx
          Local cNoCte, cTotEmails
          Local cHoraInicio
          Local xml_Link, xml_File, pdf_Link, pdf_File, folder
          Local aMailComercial, cTo, assunto, corpo, aCC := {}, aBCC := {}
          Local hMail
          Local oRow
          Local i, nLen, nEmpId, nQtdEMail
          Local nMaxMail := 9, nTotEMail := 0
          Local oEmail
          Private aEMails := {}

          //RegistraLog('oQuery:LastRec(): ' + ltrim(str(oQuery:LastRec())) )

          g_lStopExecution := .F.
          cHoraInicio      := Time()

          RegistraLog( 'Iniciando envio de e-mails de ' + LTrim(STR(oQuery:LastRec())) + ' CTEs', true )

          FOR i := 1 TO oQuery:LastRec()

              if ( g_lStopExecution )
                 MsgStatus('Interrompimento do sistema solicitado')
                 RegistraLog('Interrompimento solicitado pelo usuário, fechando o envio de e-mails...')
                 RegistraLog( 'Resumo em ' + DtoC(Date()) + ' das ' + cHoraInicio + ' às ' + Time() + ': ' + LTrim(STR(nTotEMail)) + ' e-mails enviados', true )
                 RELEASE WINDOW ALL
              end

              MsgStatus('Preparando e-mail para enviar...' + LTrim(STR(i)) + '/' + LTrim(STR(oQuery:LastRec())) + ' CTEs.' , 'emailEdit' )

              oRow := oQuery:GetRow(i)
              nEmpId := oRow:FieldGet('emp_id')

              g_oEmpresas:SetByID(nEmpId)

              if ( g_oEmpresas:nPos == 0 )
                 RegistraLog( 'ID Empresa não localizada!'     + CRLF + ;
                              'ID# : ' + LTrim(STR(nEmpId)) )
                 AADD(g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'E-Mail nao enviado. Empresa ID# ' + hb_ntos(nEmpId) + ' inexistente para este CTE!'} )
                 MsgStatus('Empresa ID# ' + LTRIM(STR(nEmpId)) + ' inexistente para o CTE!', 'emailError' )
                 LOOP
              end

              oEmail := Tsmtp_email():new(g_oEmpresas:smtp_servidor, g_oEmpresas:smtp_porta, hb_FileExists('trace_email.txt'))

              cEmitente := AllTrim(HB_ULEFT( g_oEmpresas:Nome, HB_UTF8RAT( '(', g_oEmpresas:Nome )-2 ) )
              cFone := g_oEmpresas:Fone
              nLen := HMG_LEN(HB_USUBSTR(cFone,4))
              cFone := "(" + HB_ULEFT(cFone,2) + ") " + HB_USUBSTR(cFone,4,nLen-4) + "-" + HB_URIGHT(cFone,4)
              eMailComercial := g_oEmpresas:email_comercial
              eMailContabil := g_oEmpresas:email_contabil
              cPortal := g_oEmpresas:Portal

              if ( eMailComercial == eMailContabil )
                 eMailContabil := ''
              end

              if (AllTrim(oRow:FieldGet('cte_situacao')) == 'CANCELADO')
                 pdf_Link := AllTrim(oRow:FieldGet('cte_cancelado_pdf'))
                 xml_Link := AllTrim(oRow:FieldGet('cte_cancelado_xml'))
              else
                 pdf_Link := AllTrim(oRow:FieldGet('cte_pdf'))
                 xml_Link := AllTrim(oRow:FieldGet('cte_xml'))
              end

              //https://<dominio+empresa>/mod/conhecimentos/ctes/files/35200557296543000115570010000310181000310830-cte.pdf
              pdf_File := Token(pdf_Link, '/')
              folder := g_oEmpresas:pdf_FolderDown + '20' + Substr(oRow:FieldGet('cte_chave'), 3, 4) + '\CTe\'
              if !hb_FileExists(folder + pdf_File)
               folder := 'C:\ACBrMonitorPLUS\PDF\' + SubStr(oRow:FieldGet('cte_chave'), 7, 14) + '\20' + Substr(oRow:FieldGet('cte_chave'), 3, 4) + '\CTe\'
              endif
              pdf_File := folder + pdf_File

              if hb_FileExists(pdf_File)
                 MsgStatus('Anexando PDF...' + hb_eol() + pdf_File, 'emailAttach')
                 AADD(oEmail:attachment, pdf_File)
              else
                 AADD(g_aMaiComErros, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'Arquivo PDF não encontrado no servidor local!'} )
                 MsgStatus('Arquivo PDF não encontrado!' + hb_eol() + pdf_File, 'emailError' )
                 RegistraLog('Arquivo ' + pdf_File + ' não encontrado')
              end

              //https://<dominio+empresa>/mod/conhecimentos/ctes/files/35170305197756000196570010000521271000005618.xml
              xml_File := Token(xml_Link, '/')
              folder := g_oEmpresas:xml_FolderDown + '20' + Substr(oRow:FieldGet('cte_chave'), 3, 4) + iif(AllTrim(oRow:FieldGet('cte_situacao')) == 'CANCELADO', '\Evento\Cancelamento\', '\CTe\')

              if !hb_FileExists(folder + xml_File)
               folder := 'C:\ACBrMonitorPLUS\DFes\' + SubStr(oRow:FieldGet('cte_chave'), 7, 14) + '\20' + Substr(oRow:FieldGet('cte_chave'), 3, 4) + iif(AllTrim(oRow:FieldGet('cte_situacao')) == 'CANCELADO', '\Evento\Cancelamento\', '\CTe\')
              endif
              xml_File := folder + xml_File

              if hb_FileExists(xml_File)
                 MsgStatus('Anexando XML...' + hb_eol() + xml_File, 'emailLink' )
                 AADD(oEmail:attachment, xml_File)
              else
                 AADD(g_aMaiComErros, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'Arquivo XML não encontrado no servidor local!'} )
                 MsgStatus('Arquivo XML não encontrado!' + hb_eol() + xml_File, 'emailError' )
                 RegistraLog('Arquivo ' + xml_File + ' não encontrado')
              end

              if (HMG_LEN(oEmail:attachment) == 0)
                 MsgStatus('Arquivos PDF/XML não encontrados para anexar', 'emailError' )
                 LOOP
              end

              aMailComercial := StringToVetor(eMailComercial)
              nLen           := HMG_LEN(aMailComercial)

              eMailComercial := HB_UTF8STRTRAN(eMailComercial, ",",";")
              eMailComercial := HB_UTF8STRTRAN(eMailComercial, " ")
              eMailComercial := HMG_LOWER(eMailComercial)

              if ( ';' $ eMailComercial )
                 eMailComercial := HB_ULEFT( eMailComercial, HB_UTF8AT( ';', eMailComercial ) - 1 )
              end

              oEmail:setLogin(g_oEmpresas:smtp_email, g_oEmpresas:smtp_login, g_oEmpresas:smtp_senha, g_oEmpresas:smtp_pass, eMailComercial)
              cNoCte := LTrim(Str(oRow:FieldGet('cte_numero')))

              MsgStatus('Preparando e-mail do CTE '+ cNoCte + ' (' + g_oEmpresas:sigla_cia + ')', 'emailEdit' )

              /* Verifica Ambiente Sefaz */
              if ( g_oEmpresas:Ambiente == 1 )
                 // Produção

                 assunto := 'Conhecimento de Transporte Eletrônico  (CT-e) Nº ' + cNoCte + ' ** ' + AllTrim(oRow:FieldGet('cte_situacao')) + ' **'

                 if ( oRow:FieldGet('cte_exibe_consulta_cliente') == 1 )

                    GetArrayEmails( oRow:FieldGet('clie_tomador_id'), oRow:FieldGet('clie_remetente_id'), oRow:FieldGet('clie_coleta_id'), oRow:FieldGet('clie_expedidor_id'), oRow:FieldGet('clie_recebedor_id'), oRow:FieldGet('clie_destinatario_id'), oRow:FieldGet('clie_entrega_id') )

                    if (nLen == 0)
                       AADD(g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail comercial da empresa nao cadastrado!'} )
                    end

                    if Empty(aEMails)
                       // Clientes sem e-mail
                       // Adiciona ao log de enventos para insert na tabela ctes_eventos
                       AADD(g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mails de clientes nao cadastrados!'} )
                    else
                       cTo := GetEmailTomador( oRow:FieldGet('clie_tomador_id') )
                    end

                    if Empty(cTo)

                       // Tomador sem e-mail
                       // Adiciona ao log de enventos para insert na tabela ctes_eventos

                       AADD(g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail do tomador nao cadastrado!'} )

                       if (nLen > 0)

                          cTo := aMailComercial[1]

                          if (nLen == 1)
                              aMailComercial := {}
                          else
                             aMailComercial := DelItemArray(aMailComercial, 1)
                          end

                       end

                    end

                 else

                    // Nao envia e-mail para clientes, CT-e apenas para acompanhar carga entre filiais do emitente
                    AADD(g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'CT-e apenas para acompanhar carga do Emitente, nao enviado e-mails aos clientes!'} )

                    if (nLen > 0)

                       cTo := aMailComercial[1]

                       if (nLen == 1)
                           aMailComercial := {}
                       else
                          aMailComercial := DelItemArray( aMailComercial, 1 )
                       end

                    end

                 end

                 MsgStatus('Enviando e-mail CTE '+ cNoCte + ' (' + g_oEmpresas:sigla_cia + ')', 'emailOpen' )

                 nLen := HMG_LEN(aMailComercial)

                 /* Verifica se é para enviar com cópia oculta */
                 if ( g_oEmpresas:email_CCO == 1 )

                    if (nLen > 0)
                       FOR EACH cMail IN aMailComercial
                           if ( hb_Ascan( aBCC, cMail ) == 0 )
                              AADD( aBCC, cMail )
                           end
                       NEXT EACH
                    end

                    FOR EACH hMail IN aEMails
                        if ( hb_Ascan( aBCC, hMail['email'] ) == 0 )
                           AADD( aBCC, hMail['email'] )
                        end
                    NEXT EACH

                 else

                    FOR EACH cMail IN aMailComercial
                        if ( hb_Ascan(aCC, cMail) == 0 )
                           AADD( aCC, cMail )
                        end
                    NEXT EACH

                    FOR EACH hMail IN aEMails
                        if ( hb_Ascan(aCC, hMail['email']) == 0 )
                           AADD( aCC, hMail['email'] )
                        end
                    NEXT EACH

                 end

              else

                 // Homologação
                 assunto := 'AMBIENTE DE TESTE - Conhecimento de Transporte Eletrônico  (CT-e) Nº ' + cNoCte +  + ' ** ' + AllTrim(oRow:FieldGet('cte_situacao')) + ' - EM HOMOLOGAÇÃO **'
                 cTo := aMailComercial[1]

                 /* Verifica se é para enviar com cópia oculta */
                 if ( g_oEmpresas:email_CCO == 1 )
                    aBCC := {'suporte@sistrom.com.br'}
                 else
                    aCC  := {'suporte@sistrom.com.br'}
                 end

              end

              MsgStatus('Enviando e-mail CTE '+ cNoCte + ' (' + g_oEmpresas:sigla_cia + ')', 'emailIcon' )

              if !Empty(cTo)
                 corpo := '<html>' + hb_eol() + '<body>' + hb_eol()
                 corpo += assunto + CRLF + CRLF
                 corpo += 'ENVIO DE CT-e' + CRLF + CRLF + CRLF
                 corpo += 'Esta empresa não envia SPAM! Este é um e-mail obrigatório por lei.' + CRLF + CRLF
                 corpo += 'Voce esta recebendo um Conhecimento de Transporte Eletrônico de ' + cEmitente + '.' + CRLF
                 corpo += 'Caso nao queira receber este e-mail, favor entrar em contato pelo e-mail comercial ' + eMailComercial + '.' + CRLF + CRLF
                 corpo += 'O arquivo XML do CT-e encontra-se anexado a este e-mail.' + CRLF
                 corpo += 'Para verificar a autorização do CT-e junto a SEFAZ, acesse o Portal de consulta através do endereço: https://www.cte.fazenda.gov.br.' + CRLF + CRLF
                 corpo += 'No campo "Chave de acesso", inclua a numeração da chave de acesso abaixo (sem o literal "CTe") e complete a consulta com as informações solicitadas pelo Portal.' + CRLF + CRLF + CRLF
                 corpo += 'Chave de acesso:  ' + oRow:FieldGet('cte_chave') + CRLF + CRLF + CRLF + CRLF
                 corpo += 'Atenciosamente,' + CRLF + CRLF
                 corpo += cEmitente + CRLF
                 corpo += cFone + CRLF
                 corpo += eMailComercial + CRLF + CRLF
                 corpo += 'TMS Expresso.Cloud' + CRLF

                 if !Empty(cPortal)
                    corpo += 'Acompanhe sua carga pelo portal' + CRLF
                    corpo += cPortal + CRLF
                 end

                 corpo +=  CRLF + CRLF + '*** Esse é um e-mail automático. Não é necessário respondê-lo ***' + CRLF
                 corpo += '</body>' + hb_eol() + '</html>'

                 oEmail:setRecipients(cTo, aCC, aBCC)
                 oEmail:setMsg(assunto, corpo)

                 // Controla a qtde de email enviado no dia, caso execeda a 70 emails no dia, alterna para o segundo servidor de e-Mail

                 nQtdEMail := HMG_LEN(aCC) + HMG_LEN(aBCC) + 1

                 MsgStatus('Enviando e-mails do CTE ' + cNoCte + ' (' + g_oEmpresas:sigla_cia + ')', 'emailSend' )
                 RegistraLog( 'Enviando e-mail CTE '+ cNoCte + ' (' + g_oEmpresas:sigla_cia + ') | CC: ' + oEmail:cc_as_string() + ' | BCC: ' + oEmail:bcc_as_string() + ' | nQtdEMail: ' + LTRIM(STR(nQtdEMail)) )

                 if oEmail:sendmail()

                    // Adiciona ao log de enventos para insert na tabela ctes_eventos
                    AADD( g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado para ' + oEmail:recipients['To']} )

                    FOR EACH cMail IN aCC
                        AADD(g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado com copia para ' + cMail} )
                    NEXT

                    FOR EACH cMail IN aBCC
                        AADD( g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado com copia para ' + cMail} )
                    NEXT
                    RegistraLog( 'e-Mail CTE '+ cNoCte + ' (' + g_oEmpresas:sigla_cia + ') enviado com sucesso' )
                    nTotEMail += nQtdEMail
                 else

                    MsgStatus('Falha enviando e-mail CTE: ' + cNoCte + ' (' + g_oEmpresas:sigla_cia + ')', 'emailError' )

                    AADD( g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'erro ao enviar e-mails!'} )
                    AADD( g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'Server: ' + oEmail:server + '| Porta: ' + hb_ntos(oEMail:port) + CRLF + 'De: ' + oEMail:login['From'] + CRLF + 'Para: ' + oEMail:recipients['To'] + CRLF + 'Assunto: ' + oEMail:msg['Subject']} )

                    cIPEx := IP_EXTERNO()

                    if !Empty(cIPEx)
                       AADD( g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'IP Externo: ' + cIPEx} )
                    end

                    RegistraLog( 'Email não enviado! Server: ' + oEmail:server + '| Porta: ' + hb_ntos(oEMail:port) + ' | De: ' + oEMail:login['From'] + ' | Para: ' + oEMail:recipients['To'] + ' | Assunto: ' + oEMail:msg['Subject'] )

                 end

              end

              /* Verifica novamente Ambiente Sefaz */
              if ( g_oEmpresas:Ambiente == 1 )
                 // Produção: Agora envia para os contadores só o XML em um email separado

                 MsgStatus('Enviando e-mail contabilidade', 'emailEdit' )

                 oEmail:reset()

                 aCC  := StringToVetor(eMailContabil)
                 aBCC := {}

                 aMailComercial := {}

                 nLen := HMG_LEN(aCC)

                 if (nLen == 0)
                    AADD( g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail da contabilidade do emitente nao cadastrado!'} )
                 else

                    cTo := aCC[1]

                    if (nLen == 1)
                        aCC := {}
                    else
                       aCC := DelItemArray( aCC, 1 )
                    end

                    nLen := HMG_LEN(aCC)

                    MsgStatus('Anexando XML...', 'emailAttach' )
                    AADD(oEmail:attachment, xml_File)

                    // Controla a qtde de email enviado no dia, caso execeda a 50 emails no dia, alterna para o segundo servidor de e-Mail

                    nQtdEMail := HMG_LEN(aCC) + 1

                    MsgStatus('Enviando e-mail contabilidade', 'emailSend' )
                    RegistraLog( 'Enviando e-mail para contabilidade ' + eMailContabil + ' | CTE: ' + cNoCte + ' (' + g_oEmpresas:sigla_cia + ')' )

                    if oEmail:sendmail()

                       AADD( g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado para ' + cTo} )

                       FOR EACH cMail IN aCC
                           AADD( g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado com copia para ' + cMail} )
                       NEXT

                       RegistraLog( 'e-Mail CTE '+ cNoCte + ' (' + g_oEmpresas:sigla_cia + ') para contabilidade enviado com sucesso')

                       nTotEMail += nQtdEMail

                    else
                       MsgStatus('Falha enviando e-mail CTE: ' + cNoCte + ' (' + g_oEmpresas:sigla_cia + ')', 'emailError' )
                       AADD( g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'erro ao enviar e-mail para contador!'} )
                       AADD( g_aMaiLogEvent, {'emp_id' => nEmpId, 'cte_id' => oRow:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'Server: ' + oEmail:server + '| Porta: ' + hb_ntos(oEMail:port) + CRLF + 'De: ' + oEMail:login['From'] + CRLF + 'Para: ' + oEMail:recipients['To'] + CRLF + 'Assunto: ' + oEMail:msg['Subject']} )
                    end

                 end

              end

              DO EVENTS

              if i > nMaxMail
                 // A cada 10 emails, faz update no BD
                 MsgStatus()
                 UpLoadMailEvents()
                 nMaxMail += 10
              end

              MsgStatus("Próximo envio de e-mail em 5's", 'emailIcon')
              SysWait(5, .T.)   // Pausa de 5 segundos entre cada envio de emails.

          NEXT i

          RegistraLog( 'Resumo das ' + cHoraInicio + ' às ' + Time() + ', foram enviados ' + LTrim(STR(nTotEMail)) + ' e-mails.', true )
          MsgStatus()

          g_lStopExecution := .T.

Return

Procedure GetArrayEmails( _tomador_id, _remetente_id, _coleta_id, _expedidor_id, _recebedor_id, _destinatario_id, _entrega_id )
          Local sql, ceMail
          Local oQuery, oRow
          Local j

          sql := "SELECT clie_id, con_email_cte FROM clientes_contatos WHERE clie_id IN ("
          sql += LTrim(Str(_tomador_id))

          if _remetente_id > 0 .and. !( LTrim(Str(_remetente_id))       $ LTrim(Str(_tomador_id)) + '# ' + LTrim(Str(_coleta_id))    + '# ' + LTrim(Str(_expedidor_id)) + '# ' + LTrim(Str(_recebedor_id)) + '# ' + LTrim(Str(_destinatario_id)) + '# ' + LTrim(Str(_entrega_id)) )
             sql += ", " + LTrim(Str(_remetente_id))
          end
          if _coleta_id > 0 .and. !( LTrim(Str(_coleta_id))             $ LTrim(Str(_tomador_id)) + '# ' + LTrim(Str(_remetente_id)) + '# ' + LTrim(Str(_expedidor_id)) + '# ' + LTrim(Str(_recebedor_id)) + '# ' + LTrim(Str(_destinatario_id)) + '# ' + LTrim(Str(_entrega_id)) )
             sql += ", " + LTrim(Str(_coleta_id))
          end
          if _expedidor_id > 0 .and. !( LTrim(Str(_expedidor_id))       $ LTrim(Str(_tomador_id)) + '# ' + LTrim(Str(_remetente_id)) + '# ' + LTrim(Str(_coleta_id))    + '# ' + LTrim(Str(_recebedor_id)) + '# ' + LTrim(Str(_destinatario_id)) + '# ' + LTrim(Str(_entrega_id)) )
             sql += ", " + LTrim(Str(_expedidor_id))
          end
          if _recebedor_id > 0 .and. !( LTrim(Str(_recebedor_id))       $ LTrim(Str(_tomador_id)) + '# ' + LTrim(Str(_remetente_id)) + '# ' + LTrim(Str(_coleta_id))    + '# ' + LTrim(Str(_expedidor_id)) + '# ' + LTrim(Str(_destinatario_id)) + '# ' + LTrim(Str(_entrega_id)) )
             sql += ", " + LTrim(Str(_recebedor_id))
          end
          if _destinatario_id > 0 .and. !( LTrim(Str(_destinatario_id)) $ LTrim(Str(_tomador_id)) + '# ' + LTrim(Str(_remetente_id)) + '# ' + LTrim(Str(_coleta_id))    + '# ' + LTrim(Str(_expedidor_id)) + '# ' + LTrim(Str(_recebedor_id))    + '# ' + LTrim(Str(_entrega_id)) )
             sql += ", " + LTrim(Str(_destinatario_id))
          end
          if _entrega_id > 0 .and. !( LTrim(Str(_entrega_id))           $ LTrim(Str(_tomador_id)) + '# ' + LTrim(Str(_remetente_id)) + '# ' + LTrim(Str(_coleta_id))    + '# ' + LTrim(Str(_expedidor_id)) + '# ' + LTrim(Str(_recebedor_id))    + '# ' + LTrim(Str(_destinatario_id)) )
             sql += ", " + LTrim(Str(_entrega_id))
          end

          sql += ") AND con_email_cte IS NOT NULL AND con_email_cte != '' AND con_recebe_cte != 'N' GROUP BY con_email_cte;"

          oQuery := ExecutaQuery(sql)

          if ExecutouQuery( @oQuery, sql )

             if oQuery:NetErr()
                MsgStatus('Erro SQL', 'emailError')
                PlayExclamation()
                System.Clipboard := 'Erro SQL: ' + CRLF + oQuery:Error() + CRLF + 'Descrição do comando SQL: ' + CRLF + sql  // Debug
                MsgExclamation('Erro SQL: ' + CRLF + oQuery:Error(), 'eMailCTe: Carregando Contatos')
                MsgExclamation({'Descrição do comando SQL: ', CRLF, sql, CRLF, 'Ligue para o Suporte ou tente reiniciar o programa'}, 'eMailCTe: Erro de SQL')
                oQuery:Destroy()
                RegistraLog('Carregando Contatos. Erro SQL: ' + oQuery:Error())
                RELEASE WINDOW ALL
             else


                FOR j := 1 TO oQuery:LastRec()

                    oRow   := oQuery:GetRow(j)
                    ceMail := AllTrim(oRow:FieldGet('con_email_cte'))
                    ceMail := HMG_LOWER(ceMail)

                    if ( hb_Ascan( aEMails, {|hVal| hVal['email'] == ceMail} ) == 0 )
                        AADD(aEMails, {'clie_id' => oRow:FieldGet('clie_id'), 'email' => ceMail})
                    end

                NEXT

             end

             oQuery:Destroy()

          else

             RegistraLog('Solicitação SQL ao Servidor foi perdida' + CRLF + '| SQL: ' + sql)
             MsgStatus('Solicitação SQL ao Servidor foi perdida', 'dbError')
             PlayExclamation()
             MsgExclamation({'Solicitação SQL ao Servidor foi perdida', CRLF, sql}, 'eMailCTe: Contatos')
             RELEASE WINDOW ALL

          end

Return

Function GetEmailTomador( _id )
         Local cResult := ''
         Local aNew := {}, hEmail

         FOR EACH hEmail IN aEMails

             if ( hEmail['clie_id'] == _id ) .and. Empty( cResult )
                cResult := hEmail['email']    // Achou 1º email do tomador, não adiciona este email ao array aNew
             else
                AAdd( aNew, hb_HClone(hEmail) )
             end

         NEXT EACH

         //AEVAL( aEMails, { |hE | if( hE['clie_id'] == _id, cResult := hE['email'], AADD( aNew, hE ) ) } )

         // Clona o novo array sem o tomador

         if !Empty( cResult )
            aEMails := AClone( aNew )
         end

         RegistraLog('EMail Tomador (cResult): ' + cResult + ' | aEMails: ' + LTRIM(STR(hmg_len(aEMails))) )

Return ( cResult )

Function StringToVetor(cString)
         Local p
         Local cMail
         Local aVetor := {}

         if !Empty(cString)

            cString := HB_UTF8STRTRAN(cString, ",",";")
            cString := HB_UTF8STRTRAN(cString, " ")
            cString := HMG_LOWER(cString)

            While (p := HB_UTF8AT(';', cString)) > 0

                  cMail := HB_ULEFT(cString, p-1)

                  if ( hb_Ascan( aVetor,  cMail ) == 0 )
                     AADD(aVetor, cMail)
                  end

                  cString := HB_USUBSTR(cString, p+1)

            EndDo

            if ( hb_Ascan( aVetor,  cString ) == 0 )
               AADD(aVetor, cString)
            end

         end

Return AClone(aVetor)

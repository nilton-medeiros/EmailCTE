#include <hmg.ch>
#include "hbclass.ch"

#define true  .T.
#define false .F.
#define SKIP_LINE .T.
#define NO_SKIP_LINE .F.
#define ENCRYPTED .T.

procedure send_emails(ctes)
    local oEmail, emails_comerciais, envios, cte, empresa := g_oEmpresas, prepare
    local len, emp_id, nQtdEmail, nMaxMail := 9, total_emails := 0
    local foneFormatted, ctePath, cte_numero, assunto, start_time
    local pdf_Link, pdf_file, xml_Link, xml_File, ip_externo, emails_string
    local emitente := {;
            "nome" => '',;
            "fone" => '',;
            "emailComercial" => '',;
            "emailContabil" => '',;
            "portal" => '';
          }
    local cTo := '', cc := {}, bcc := {}
    local hEmail, cMail

    registraLog('Iniciando envio de e-mails de ' + hb_ntos(ctes:LastRec()) + ' CTEs', SKIP_LINE)

    g_lStopExecution := false
    start_time := Time()

    for i := 1 to ctes:LastRec()

        // g_lStopExecution: Variável pública, o sistema pode ser interrompido ou não em um processo de loop
        if ( g_lStopExecution )
            MsgStatus('Interrompimento do sistema solicitado')
            registraLog('Interrompimento solicitado pelo usuário, fechando o envio de e-mails...')
            registraLog('Resumo em ' + DtoC(Date()) + ' das ' + start_time + ' às ' + Time() + ': ' + hb_ntos(total_emails) + ' e-mails enviados', SKIP_LINE)
            RELEASE WINDOW ALL
        end

        MsgStatus('Preparando e-mail para enviar...' + hb_ntos(i) + '/' + hb_ntos(ctes:LastRec()) + ' CTEs.' , 'emailEdit')

        cte := ctes:GetRow(i)
        emp_id := cte:FieldGet('emp_id')
        empresa:SetByID(emp_id)

        if (empresa:id == 0)

            registraLog('ID Empresa não localizada!' + hb_eol() + ;
                'ID# : ' + hb_ntos(emp_id);
            )

            AAdd(g_aMaiLogEvent,;
                {;
                    "emp_id" => emp_id,;
                    "cte_id" => cte:FieldGet("cte_id"),;
                    "data" => date(),;
                    "hora" => time(),;
                    "mensagem" => 'E-Mail nao enviado. Empresa ID# ' + hb_ntos(emp_id) + ' inexistente para este CTE!';
                };
            )

            MsgStatus('Empresa ID# ' + hb_ntos(emp_id) + ' inexistente para o CTE!', 'emailError')

            loop

        endif

        foneFormatted := empresa:Fone
        len := hmg_len(hb_USubStr(foneFormatted, 4))
        foneFormatted := "(" + hb_ULeft(foneFormatted, 2) + ") " + hb_USubStr(foneFormatted, 4, len-4) + "-" + hb_URight(foneFormatted, 4)

        emitente["fone"] := foneFormatted
        emitente["nome"] := AllTrim(hb_ULeft(empresa:Nome, hb_utf8RAt('(', empresa:Nome) - 2))
        emitente["portal"] := empresa:Portal
        emitente["emailComercial"] := empresa:email_comercial

        if !(empresa:email_comercial == empresa:email_contabil)
            emitente["emailContabil"] := empresa:email_contabil
        endif

        if (AllTrim(cte:FieldGet("cte_situacao")) == 'CANCELADO')
            pdf_Link := AllTrim(cte:FieldGet("cte_cancelado_pdf"))
            xml_Link := AllTrim(cte:FieldGet("cte_cancelado_xml"))
        else
            pdf_Link := AllTrim(cte:FieldGet("cte_pdf"))
            xml_Link := AllTrim(cte:FieldGet("cte_xml"))
        endif

        //https://www.<site>.com.br/<agente>/mod/conhecimentos/ctes/files/35200557296543000115570010000310181000310830-cte.pdf

        ctePath := empresa:cte_path + '20' + Substr(cte:FieldGet('cte_chave'), 3, 4) + "\"
        pdf_file := Token(pdf_Link, '/')

        if !hb_FileExists(ctePath + pdf_file)
            ctePath := RegistryRead("HKEY_CURRENT_USER\SOFTWARE\Sistrom\SendToPrinter\InstallPath")
            if !Empty(ctePath)
                ctePath += "printed\"
            endif
        endif

        oEmail := Tsmtp_email():new(empresa:smtp_servidor, empresa:smtp_porta, hb_FileExists('trace_email.txt'))

        if hb_FileExists(ctePath + pdf_file)
            MsgStatus('Anexando PDF...' + hb_eol() + pdf_file, 'emailAttach')
            oEmail:attachFile(ctePath + pdf_file)
        else
            AAdd(g_aMaiComErros, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'Arquivo PDF não encontrado no servidor local!'})
            MsgStatus('Arquivo PDF não encontrado!' + hb_eol() + pdf_file, 'emailError' )
            registraLog('Arquivo ' + ctePath + pdf_file + ' não encontrado')
        end

        //https://www.<site>/<agente>/mod/conhecimentos/ctes/files/35230857296543000115570010000441261000631947-cte.xml

        xml_File := Token(xml_Link, '/')    // 35230857296543000115570010000441261000631947-cte.xml

        if hb_FileExists(ctePath + xml_File)
            MsgStatus('Anexando XML...' + hb_eol() + xml_File, 'emailLink' )
            oEmail:attachFile(ctePath + xml_File)
        else
            AAdd(g_aMaiComErros, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'Arquivo XML não encontrado no servidor local!'})
            MsgStatus('Arquivo XML não encontrado!' + hb_eol() + xml_File, 'emailError' )
            registraLog('Arquivo ' + ctePath + xml_File + ' não encontrado')
        end

        if oEmail:is_not_attached()
            MsgStatus('Arquivos PDF/XML não encontrados para anexar', 'emailError' )
            LOOP
        end

        cte_numero := hb_ntos(cte:FieldGet("cte_numero"))

        MsgStatus( 'Preparando e-mail do CTe ' + cte_numero + ' (' + empresa:sigla_cia + ')', 'emailEdit')

        emails_comerciais := TEmailsList():new(emitente["emailComercial"])
        assunto := 'Conhecimento de Transporte Eletrônico  (CT-e) Nº ' + cte_numero + ' ** ' + AllTrim(cte:FieldGet('cte_situacao')) + ' **'

        oEmail:setLogin(empresa:smtp_email, empresa:smtp_login, empresa:smtp_senha, empresa:smtp_pass, emails_comerciais:getTo())

        if (empresa:Ambiente == 1)
            // Produção
            if emails_comerciais:is_email()
                cTo := emails_comerciais:getTo()
            else
                cTo := ''
                AAdd(g_aMaiLogEvent, {;
                                        'emp_id' => emp_id,;
                                        'cte_id' => cte:FieldGet('cte_id'),;
                                        'data' => date(),;
                                        'hora' => time(),;
                                        'mensagem' => 'e-mail comercial da empresa emitente nao cadastrado!';
                                     })
            endif

            if is_true(cte:FieldGet("enviar_email_cliente"))
                envios := nil
                envios := TEmails():new(;
                            cte:FieldGet('clie_tomador_id'),;
                            cte:FieldGet('clie_remetente_id'),;
                            cte:FieldGet('clie_coleta_id'),;
                            cte:FieldGet('clie_expedidor_id'),;
                            cte:FieldGet('clie_recebedor_id'),;
                            cte:FieldGet('clie_destinatario_id'),;
                            cte:FieldGet('clie_entrega_id'))

                if !envios:temTomador
                    // Tomador sem e-mail
                    // Adiciona ao log de enventos para insert na tabela ctes_eventos
                    AAdd(g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail do tomador nao cadastrado!'})
                endif

                if Empty(envios:getEmailTo())
                    // Clientes sem e-mail
                    // Adiciona ao log de enventos para insert na tabela ctes_eventos
                    AAdd(g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mails de clientes nao cadastrados!'})
                else
                    cTo := envios:getEmailTo()
                endif

                if is_true(empresa:email_CCO)
                    // Enviar com cópia oculta
                    for each cMail in emails_comerciais:emails
                        AAdd(bcc, cMail)
                    next
                    for each hEmail in envios:emails
                        AAdd(bcc, hEmail["email"])
                    next
                else
                    // Enviar com cópia normal
                    for each cMail in emails_comerciais:emails
                        AAdd(cc, cMail)
                    next
                    for each hEmail in envios:emails
                        AAdd(cc, hEmail["email"])
                    next
                endif
                envios:destroy()

            else
                // Nao envia e-mail para clientes, CT-e apenas para acompanhar carga entre filiais do emitente
                AAdd(g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'CT-e apenas para acompanhar carga do Emitente, nao enviado e-mails aos clientes!'} )
            endif

            oEmail:setRecipients(cTo, cc, bcc)

        else
            // Homologação
            assunto := 'AMBIENTE DE TESTE - ' + assunto + ' - EM HOMOLOGAÇÃO **'
            cTo := ifNull(emails_comerciais:getTo(), 'suporte@sistrom.com.br')

            /* Verifica se é para enviar com cópia oculta */
            if is_true(empresa:email_CCO)
                oEmail:setRecipients(cTo, {}, {'suporte@sistrom.com.br'})
            else
                oEmail:setRecipients(emails_comerciais:getTo(), {'suporte@sistrom.com.br'}, {})
            end
        endif

        MsgStatus('Enviando e-mail CTE '+ cte_numero + ' (' + empresa:sigla_cia + ')', 'emailOpen' )

        prepare := {=>}
        prepare["assunto"] := assunto
        prepare["cteChave"] := cte:FieldGet('cte_chave')
        prepare["nomeRemetente"] := emitente['nome']
        prepare["foneRemetente"] := emitente['fone']
        prepare["portal"] := emitente['portal']

        oEmail:prepare(prepare)

        nQtdEmail := hmg_len(cc) + hmg_len(bcc) + 1

        MsgStatus('Enviando e-mails do CTE ' + cte_numero + ' (' + empresa:sigla_cia + ')', 'emailSend' )
        registraLog( 'Enviando e-mail CTE '+ cte_numero + ' (' + empresa:sigla_cia + ') | CC: ' + oEmail:cc_as_string() + ' | BCC: ' + oEmail:bcc_as_string() + ' | nQtdEMail: ' + hb_ntos(nQtdEMail))

        if oEmail:sendmail()

            // Adiciona ao log de enventos para insert na tabela ctes_eventos
            AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado para ' + oEmail:recipients['To']})

            FOR EACH cMail IN cc
               AAdd(g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado com copia para ' + cMail})
            NEXT

            FOR EACH cMail IN bcc
               AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado com copia para ' + cMail})
            NEXT
            registraLog( 'e-Mail CTE '+ cte_numero + ' (' + empresa:sigla_cia + ') enviado com sucesso')
            total_emails += nQtdEMail

            if Empty(emitente["emailContabil"])
                AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail da contabilidade do emitente nao cadastrado!'} )
            else

                // Enviar emails para Contabilidade/Contadores somente o XML em um email separado
                MsgStatus('Enviando e-mail contabilidade', 'emailEdit' )

                oEmail:reset()

                emails_string := hb_utf8StrTran(emitente["emailContabil"], ",", ";")
                emails_string := hb_utf8StrTran(emails_string, " ")
                emails_string := hmg_lower(emails_string)
                cc := hb_ATokens(emails_string, ";")
                bcc := {}
                cTo := cc[1]
                hb_adel(cc, 1, true)

                oEmail:setRecipients(cTo, cc, bcc)

                MsgStatus('Anexando XML...', 'emailAttach' )
                oEmail:attachFile(ctePath + xml_File)

                // Controla a qtde de email enviado no dia, caso execeda a 50 emails no dia, alterna para o segundo servidor de e-Mail
                nQtdEMail := HMG_LEN(cc) + 1

                MsgStatus('Enviando e-mail contabilidade', 'emailSend' )
                registraLog( 'Enviando e-mail para contabilidade ' + emails_string + ' | CTE: ' + cte_numero + ' (' + empresa:sigla_cia + ')')

                if oEmail:sendmail()

                    AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado para ' + cTo})

                    FOR EACH cMail IN cc
                        AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado com copia para ' + cMail} )
                    NEXT

                    registraLog( 'e-Mail CTE '+ cte_numero + ' (' + empresa:sigla_cia + ') para contabilidade enviado com sucesso')

                    total_emails += nQtdEMail

                 else
                    MsgStatus('Falha enviando e-mail CTE: ' + cte_numero + ' (' + empresa:sigla_cia + ')', 'emailError' )
                    AAdd(g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'erro ao enviar e-mail para contador!'} )
                    AAdd(g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'Server: ' + oEmail:server + '| Porta: ' + hb_ntos(oEmail:port) + CRLF + 'De: ' + oEmail:login['From'] + CRLF + 'Para: ' + oEmail:recipients['To'] + CRLF + 'Assunto: ' + oEmail:msg['Subject']})
                endif

           endif

        else

           MsgStatus('Falha enviando e-mail CTE: ' + cte_numero + ' (' + empresa:sigla_cia + ')', 'emailError' )

           AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'erro ao enviar e-mails!'})
           AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'Server: ' + oEmail:server + '| Porta: ' + hb_ntos(oEmail:port) + CRLF + 'De: ' + oEmail:login['From'] + CRLF + 'Para: ' + oEmail:recipients['To'] + CRLF + 'Assunto: ' + oEmail:msg['Subject']})

           ip_externo := IP_Externo()

           if !Empty(ip_externo)
              AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'IP Externo: ' + ip_externo})
           end

           registraLog( 'Email não enviado! Server: ' + oEmail:server + '| Porta: ' + hb_ntos(oEmail:port) + ' | De: ' + oEmail:login['From'] + ' | Para: ' + oEmail:recipients['To'] + ' | Assunto: ' + oEmail:msg['Subject'] )

        endif

        DO EVENTS

        if i > nMaxMail
            upload_mail_events()
            nMaxMail += 10
        endif

        MsgStatus("Próximo envio de e-mail em 5's", 'emailIcon')
        SysWait(5, .T.)   // Pausa de 5 segundos entre cada envio de emails.
        emails_comerciais:destroy()
        emails_comerciais := nil
        oEmail:destroy()
        oEmail := nil     // Destroi objeto anterior, estamos em um loop do For

    next i

    registraLog( 'Resumo das ' + start_time + ' às ' + Time() + ', foram enviados ' + hb_ntos(total_emails) + ' e-mails.', SKIP_LINE)
    MsgStatus()

    g_lStopExecution := .T.

return

/* Class ================================================================================================= */

create class TEmails

    data destinatarios INIT {{"destinatario" => "remetente", "enviar" => false, "id" => ''},;
                             {"destinatario" => "coleta", "enviar" => false, "id" => ''},;
                             {"destinatario" => "expedidor", "enviar" => false, "id" => ''},;
                             {"destinatario" => "recebedor", "enviar" => false, "id" => ''},;
                             {"destinatario" => "destinatario", "enviar" => false, "id" => ''},;
                             {"destinatario" => "entrega", "enviar" => false, "id" => ''};
                            } protected
    data emails INIT {}
    data emailTo INIT '' protected
    data temTomador INIT false

    method getEmailTo() INLINE ::emailTo
    method is_email() INLINE !(hmg_len(::emails) == 0)
    method new(tomador_id, remetente_id, coleta_id, expedidor_id, recebedor_id, destinatario_id, entrega_id) constructor
    method destroy()

end class

method new(tomador_id, remetente_id, coleta_id, expedidor_id, recebedor_id, destinatario_id, entrega_id) class TEmails
    local destinatario, contindo, query, contato, contato_email, idx
    local sql := "SELECT clie_id, con_email_cte FROM clientes_contatos WHERE clie_id IN ("

    sql += hb_ntos(tomador_id)

    for each destinatario in ::destinatarios
        switch destinatario["destinatario"]
            case 'remetente'
                destinatario["id"] := hb_ntos(remetente_id)
                destinatario["enviar"] := is_true(remetente_id)
                contido := hb_ntos(tomador_id) + '# ' + hb_ntos(coleta_id) + '# ' + hb_ntos(expedidor_id) + '# ' + hb_ntos(recebedor_id) + '# ' + hb_ntos(destinatario_id) + '# ' + hb_ntos(entrega_id)
                exit
            case 'coleta'
                destinatario["id"] := hb_ntos(coleta_id)
                destinatario["enviar"] := is_true(coleta_id)
                contido := hb_ntos(tomador_id) + '# ' + hb_ntos(remetente_id) + '# ' + hb_ntos(expedidor_id) + '# ' + hb_ntos(recebedor_id) + '# ' + hb_ntos(destinatario_id) + '# ' + hb_ntos(entrega_id)
                exit
            case 'expedidor'
                destinatario["id"] := hb_ntos(expedidor_id)
                destinatario["enviar"] := is_true(expedidor_id)
                contido := hb_ntos(tomador_id) + '# ' + hb_ntos(remetente_id) + '# ' + hb_ntos(coleta_id) + '# ' + hb_ntos(recebedor_id) + '# ' + hb_ntos(destinatario_id) + '# ' + hb_ntos(entrega_id)
                exit
            case 'recebedor'
                destinatario["id"] := hb_ntos(recebedor_id)
                destinatario["enviar"] := is_true(recebedor_id)
                contido := hb_ntos(tomador_id) + '# ' + hb_ntos(remetente_id) + '# ' + hb_ntos(coleta_id) + '# ' + hb_ntos(expedidor_id) + '# ' + hb_ntos(destinatario_id) + '# ' + hb_ntos(entrega_id)
                exit
            case 'destinatario'
                destinatario["id"] := hb_ntos(destinatario_id)
                destinatario["enviar"] := is_true(destinatario_id)
                contido := hb_ntos(tomador_id) + '# ' + hb_ntos(remetente_id) + '# ' + hb_ntos(coleta_id) + '# ' + hb_ntos(expedidor_id) + '# ' + hb_ntos(recebedor_id) + '# ' + hb_ntos(entrega_id)
                exit
            case 'entrega'
                destinatario["id"] := hb_ntos(entrega_id)
                destinatario["enviar"] := is_true(entrega_id)
                contido := hb_ntos(tomador_id) + '# ' + hb_ntos(remetente_id) + '# ' + hb_ntos(coleta_id) + '# ' + hb_ntos(expedidor_id) + '# ' + hb_ntos(recebedor_id) + '# ' + hb_ntos(destinatario_id)
                exit
        endswitch

        if destinatario["enviar"]
            destinatario["enviar"] := !(destinatario["id"] $ contido)
            if destinatario["enviar"]
                sql += ", " + destinatario["id"]
            endif
        endif

    next

    sql += ") AND con_email_cte IS NOT NULL AND con_email_cte != '' AND con_recebe_cte != 'N' GROUP BY con_email_cte;"

    query := ExecutaQuery(sql)

    if ExecutouQuery(@query, sql)
        if query:NetErr()
            MsgStatus('Erro SQL', 'emailError')
            PlayExclamation()
            //System.Clipboard := 'Erro SQL: ' + hb_eol() + query:Error() + hb_eol() + 'Descrição do comando SQL: ' + hb_eol() + sql  // Debug
            MsgExclamation('Erro SQL: ' + hb_eol() + query:Error(), 'eMailCTe: Carregando Contatos')
            MsgExclamation('Ligue para o Suporte ou tente reiniciar o programa', 'eMailCTe: Erro de SQL')
            query:Destroy()
            registraLog('Carregando Contatos. Erro SQL: ' + query:Error() + hb_eol() + 'SQL: ' + sql,, true)
            RELEASE WINDOW ALL
        else
            for j := 1 TO query:LastRec()
                contato := query:GetRow(j)
                contato_email := AllTrim(contato:FieldGet('con_email_cte'))
                contato_email := hmg_lower(contato_email)
                if (hb_Ascan( ::emails, {|hVal| hVal['email'] == contato_email} ) == 0)
                    AAdd(::emails, {'clie_id' => contato:FieldGet('clie_id'), 'email' => contato_email})
                end
            next j
        endif
        query:Destroy()
    else
        registraLog('Solicitação SQL ao Servidor foi perdida' + hb_eol() + '| SQL: ' + sql,, true)
        MsgStatus('Solicitação SQL ao Servidor foi perdida', 'dbError')
        PlayExclamation()
        MsgExclamation({'Solicitação SQL ao Servidor foi perdida', hb_eol(), sql}, 'eMailCTe: Contatos')
        RELEASE WINDOW ALL
    endif

    ::emailTo := ''

    if ::is_email()
        idx := hb_Ascan( ::emails, {|h| h["clie_id"] == tomador_id})
        ::temTomador := !(idx == 0)
        idx := iif((idx == 0), 1, idx)
        ::emailTo := ::emails[idx]["email"]
        hb_adel(::emails, idx, true)
    endif

return self

method destroy() class TEmails
    ::destinatarios := nil
    ::emails := nil
    ::emailTo := ''
return self
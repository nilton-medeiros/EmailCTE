#include <hmg.ch>
#include "hbclass.ch"

#define true  .T.
#define false .F.
#define SKIP_LINE .T.
#define NO_SKIP_LINE .F.
#define ENCRYPTED .T.

procedure send_emails(ctes)
    local email, emails_comerciais, envios, cte, empresa := g_oEmpresas
    local len, emp_id, nQtdEmail, nMaxMail := 9, total_emails := 0
    local foneFormatted, path, cte_numero, assunto, start_time
    local pdf_Link, pdf_file, xml_Link, xml_File, ip_externo, emails_string
    local emitente := {;
            "nome" => '',;
            "fone" => '',;
            "emailComercial" => '',;
            "emailContabil" => '',;
            "portal" => '',;
          }
    local cTo := '', cc := {}, bcc := {}

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
        foneFormatted := "(" + hb_ULeft(foneFormatted, 2) + ") " + hb_USubStr(foneFormatted, 4, ten-4) + "-" + hb_URight(foneFormatted, 4)

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
        pdf_file := Token(pdf_Link, '/')
        path := empresa:pdf_FolderDown + '20' + Substr(cte:FieldGet('cte_chave'), 3, 4) + '\CTe\'

        if !hb_FileExists(path + pdf_file)
            path := 'C:\ACBrMonitorPLUS\PDF\' + SubStr(cte:FieldGet('cte_chave'), 7, 14) + '\20' + Substr(cte:FieldGet('cte_chave'), 3, 4) + '\CTe\'
        endif

        pdf_file := path + pdf_file

        email := Tsmtp_email():new(empresa:smtp_servidor, empresa:smtp_porta, hb_FileExists('trace_email.txt'))

        if hb_FileExists(pdf_file)
            MsgStatus('Anexando PDF...' + hb_eol() + pdf_file, 'emailAttach')
            email:attachFile(pdf_file)
        else
            AAdd(g_aMaiComErros, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'Arquivo PDF não encontrado no servidor local!'})
            MsgStatus('Arquivo PDF não encontrado!' + hb_eol() + pdf_file, 'emailError' )
            registraLog('Arquivo ' + pdf_file + ' não encontrado')
        end

        //https://www.<site>/<agente>/mod/conhecimentos/ctes/files/35170305197756000196570010000521271000005618.xml
        xml_File := Token(xml_Link, '/')
        path := empresa:xml_FolderDown + '20' + Substr(cte:FieldGet('cte_chave'), 3, 4) + iif(AllTrim(cte:FieldGet('cte_situacao')) == 'CANCELADO', '\Evento\Cancelamento\', '\CTe\')

        if !hb_FileExists(path + xml_File)
            path := 'C:\ACBrMonitorPLUS\DFes\' + SubStr(cte:FieldGet('cte_chave'), 7, 14) + '\20' + Substr(cte:FieldGet('cte_chave'), 3, 4) + iif(AllTrim(cte:FieldGet('cte_situacao')) == 'CANCELADO', '\Evento\Cancelamento\', '\CTe\')
        endif
        xml_File := path + xml_File

        if hb_FileExists(xml_File)
            MsgStatus('Anexando XML...' + hb_eol() + xml_File, 'emailLink' )
            email:attachFile(xml_File)
        else
            AAdd(g_aMaiComErros, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'Arquivo XML não encontrado no servidor local!'})
            MsgStatus('Arquivo XML não encontrado!' + hb_eol() + xml_File, 'emailError' )
            registraLog('Arquivo ' + xml_File + ' não encontrado')
        end

        if email:is_not_attached()
            MsgStatus('Arquivos PDF/XML não encontrados para anexar', 'emailError' )
            LOOP
        end

        MsgStatus('Preparando e-mail do CTe ' + cte_numero + ' (' + empresa:sigla_cia + ')', 'emailEdit')

        emails_comerciais := TEmailsList():new(emitente["emailComercial"])
        cte_numero := hb_ntos(cte:FieldGet("cte_numero"))
        assunto := 'Conhecimento de Transporte Eletrônico  (CT-e) Nº ' + cte_numero + ' ** ' + AllTrim(cte:FieldGet('cte_situacao')) + ' **'

        email:setLogin(empresa:smtp_email, empresa:smpt_login, empresa:senha, empresa:pass, emails_comerciais:getTo())

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
                    for each email in emails_comerciais:emails
                        AAdd(bcc, email)
                    next
                    for each hEmail in envios:emails
                        AAdd(bcc, hEmail["email"])
                    next
                else
                    // Enviar com cópia normal
                    for each email in emails_comerciais:emails
                        AAdd(cc, email)
                    next
                    for each hEmail in envios:emails
                        AAdd(cc, hEmail["email"])
                    next
                endif
                envios:destroy()
                email:setRecipients(cTo, cc, bcc)

            else
                // Nao envia e-mail para clientes, CT-e apenas para acompanhar carga entre filiais do emitente
                AAdd(g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'CT-e apenas para acompanhar carga do Emitente, nao enviado e-mails aos clientes!'} )
                email:setRecipients(cTo, cc, bcc)
            endif

        else
            // Homologação
            assunto := 'AMBIENTE DE TESTE - ' + assunto + ' - EM HOMOLOGAÇÃO **'
            cTo := ifNull(emails_comerciais:getTo(), 'suporte@sistrom.com.br')

            /* Verifica se é para enviar com cópia oculta */
            if is_true(empresa:email_CCO)
                email:setRecipients(cTo, {}, {'suporte@sistrom.com.br'})
            else
                email:setRecipients(emails_comerciais:getTo(), {'suporte@sistrom.com.br'}, {})
            end
        endif

        MsgStatus('Enviando e-mail CTE '+ cte_numero + ' (' + empresa:sigla_cia + ')', 'emailOpen' )

        email:prepare(;
            {;
                "assunto" => assunto,;
                "cteChave" => cte:FieldGet('cte_chave'),;
                "nomeRemetente" => emitente['nome'],;
                "foneRemetente" => emitente['fone'],;
                "portal" => emitente['portal'];
            })

        nQtdEmail := hmg_len(cc) + hmg_len(bcc) + 1

        MsgStatus('Enviando e-mails do CTE ' + cte_numero + ' (' + empresa:sigla_cia + ')', 'emailSend' )
        registraLog( 'Enviando e-mail CTE '+ cte_numero + ' (' + empresa:sigla_cia + ') | CC: ' + email:cc_as_string() + ' | BCC: ' + email:bcc_as_string() + ' | nQtdEMail: ' + hb_ntos(nQtdEMail))

        if email:sendmail()

            // Adiciona ao log de enventos para insert na tabela ctes_eventos
            AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado para ' + email:recipients['To']})

            FOR EACH cMail IN cc
               AAdd(g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado com copia para ' + cMail})
            NEXT

            FOR EACH cMail IN bcc
               AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado com copia para ' + cMail})
            NEXT
            registraLog( 'e-Mail CTE '+ cte_numero + ' (' + empresa:sigla_cia + ') enviado com sucesso')
            nTotEMail += nQtdEMail

            if Empty(emitente["emailContabil"])
                AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail da contabilidade do emitente nao cadastrado!'} )
            else

                // Enviar emails para Contabilidade/Contadores somente o XML em um email separado
                MsgStatus('Enviando e-mail contabilidade', 'emailEdit' )

                email:reset()

                emails_string := hb_utf8StrTran(emitente["emailContabil"], ",", ";")
                emails_string := hb_utf8StrTran(emails_string, " ")
                emails_string := hmg_low(emails_string)
                cc := hb_ATokens(emails_string, ";")
                bcc := {}
                cTo := cc[1]
                hb_adel(cc, 1, true)

                email:setRecipients(cTo, cc, bcc)

                MsgStatus('Anexando XML...', 'emailAttach' )
                email:attachFile(xml_File)

                // Controla a qtde de email enviado no dia, caso execeda a 50 emails no dia, alterna para o segundo servidor de e-Mail
                nQtdEMail := HMG_LEN(cc) + 1

                MsgStatus('Enviando e-mail contabilidade', 'emailSend' )
                registraLog( 'Enviando e-mail para contabilidade ' + emails_string + ' | CTE: ' + cte_numero + ' (' + empresa:sigla_cia + ')')

                if email:sendmail()

                    AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado para ' + cTo})

                    FOR EACH cMail IN cc
                        AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'e-mail enviado com copia para ' + cMail} )
                    NEXT

                    registraLog( 'e-Mail CTE '+ cte_numero + ' (' + empresa:sigla_cia + ') para contabilidade enviado com sucesso')

                    nTotEMail += nQtdEMail

                 else
                    MsgStatus('Falha enviando e-mail CTE: ' + cte_numero + ' (' + empresa:sigla_cia + ')', 'emailError' )
                    AAdd(g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'erro ao enviar e-mail para contador!'} )
                    AAdd(g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'Server: ' + email:server + '| Porta: ' + hb_ntos(email:port) + CRLF + 'De: ' + email:login['From'] + CRLF + 'Para: ' + email:recipients['To'] + CRLF + 'Assunto: ' + email:msg['Subject']})
                endif

           endif

        else

           MsgStatus('Falha enviando e-mail CTE: ' + cte_numero + ' (' + empresa:sigla_cia + ')', 'emailError' )

           AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'erro ao enviar e-mails!'})
           AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'Server: ' + email:server + '| Porta: ' + hb_ntos(email:port) + CRLF + 'De: ' + email:login['From'] + CRLF + 'Para: ' + email:recipients['To'] + CRLF + 'Assunto: ' + email:msg['Subject']})

           ip_externo := IP_EXTERNO()

           if !Empty(ip_externo)
              AAdd( g_aMaiLogEvent, {'emp_id' => emp_id, 'cte_id' => cte:FieldGet('cte_id'), 'data' => date(), 'hora' => time(), 'mensagem' => 'IP Externo: ' + ip_externo})
           end

           registraLog( 'Email não enviado! Server: ' + email:server + '| Porta: ' + hb_ntos(email:port) + ' | De: ' + email:login['From'] + ' | Para: ' + email:recipients['To'] + ' | Assunto: ' + email:msg['Subject'] )

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
        email:destroy()
        email := nil     // Destroi objeto anterior, estamos em um loop do For

    next i

    registraLog( 'Resumo das ' + start_time + ' às ' + Time() + ', foram enviados ' + hb_ntos(nTotEMail) + ' e-mails.', SKIP_LINE)
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
                contido := hb_ntos(tomador_id) + '# ' + hb_ntos(coleta_id) + '# ' + hb_ntos(expedidor_id) + '# ' + hb_ntos(recebedor_id) + '# ' + hb_ntos(destinatario_id) + '# ' + hb_ntos(entrega_id))
                exit
            case 'coleta'
                destinatario["id"] := hb_ntos(coleta_id)
                destinatario["enviar"] := is_true(coleta_id)
                contido := hb_ntos(tomador_id) + '# ' + hb_ntos(remetente_id) + '# ' + hb_ntos(expedidor_id) + '# ' + hb_ntos(recebedor_id) + '# ' + hb_ntos(destinatario_id) + '# ' + hb_ntos(entrega_id))
                exit
            case 'expedidor'
                destinatario["id"] := hb_ntos(expedidor_id)
                destinatario["enviar"] := is_true(expedidor_id)
                contido := hb_ntos(tomador_id) + '# ' + hb_ntos(remetente_id) + '# ' + hb_ntos(coleta_id) + '# ' + hb_ntos(recebedor_id) + '# ' + hb_ntos(destinatario_id) + '# ' + hb_ntos(entrega_id))
                exit
            case 'recebedor'
                destinatario["id"] := hb_ntos(recebedor_id)
                destinatario["enviar"] := is_true(recebedor_id)
                contido := hb_ntos(tomador_id) + '# ' + hb_ntos(remetente_id) + '# ' + hb_ntos(coleta_id) + '# ' + hb_ntos(expedidor_id) + '# ' + hb_ntos(destinatario_id) + '# ' + hb_ntos(entrega_id))
                exit
            case 'destinatario'
                destinatario["id"] := hb_ntos(destinatario_id)
                destinatario["enviar"] := is_true(destinatario_id)
                contido := hb_ntos(tomador_id) + '# ' + hb_ntos(remetente_id) + '# ' + hb_ntos(coleta_id) + '# ' + hb_ntos(expedidor_id) + '# ' + hb_ntos(recebedor_id) + '# ' + hb_ntos(entrega_id))
                exit
            case 'entrega'
                destinatario["id"] := hb_ntos(entrega_id)
                destinatario["enviar"] := is_true(entrega_id)
                contido := hb_ntos(tomador_id) + '# ' + hb_ntos(remetente_id) + '# ' + hb_ntos(coleta_id) + '# ' + hb_ntos(expedidor_id) + '# ' + hb_ntos(recebedor_id) + '# ' + hb_ntos(destinatario_id))
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
        idx := hb_Ascan( ::email, {|h| h["clie_id"] == tomador_id})
        if (idx == 0)
            ::temTomador := false
        endif
        idx := iif((idx == 0), 1, idx)
        ::emailTo := ::email[idx]["email"])
        hb_adel(::email, idx, true)
    endif

return self

method destroy() class TEmails
    ::destinatarios := nil
    ::emails := nil
    ::emailTo := ''
return self
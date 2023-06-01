#include <hmg.ch>
#include "hbclass.ch"
//#include <hbextern.ch>

//#require "hbssl"
//#require "hbtip"

//#include "simpleio.ch"

#define true  .T.
#define false .F.


create class Tsmtp_email
   data server readonly
   data port readonly
   data recipients readonly
   data subject readonly
   data body readonly
   data msg readonly
   data attachment
   data login readonly
   data popServer readonly
   data priority readonly
   data isRead readonly
   data trace readonly
   data popAuth readonly
   data noAuth readonly
   data timeOut readonly
   data isTLS readonly
   data charset readonly
   data encoding readonly

   method new(server, port, trace) constructor
   method setLogin(from, user, pass, smtp_pass, replyTo)
   method setRecipients(eTo, cc, bcc)
   method prepare(content)
   method cc_as_string()
   method bcc_as_string()
   method attachFile(file)
   method is_not_attached()
   method sendmail()
   method reset()
   method destroy()
end class

method new(server, port, trace) class Tsmtp_email
   default trace := false
   ::server := server
   ::port := port
   ::msg := {'Subject' => '', 'Body' => ''}
   ::attachment := {}
   ::login := {'From' => '', 'User' => '', 'Pass' => '', 'ReplyTo' => nil, 'SMTPPass' => ''}
   ::popServer := ''
   ::trace := trace
   ::popAuth := false
   ::noAuth := false
   ::isTLS := (::port == 465)
return self

method setLogin(from, user, pass, smtp_pass, replyTo) class Tsmtp_email
   ::login['From'] := from
   ::login['User'] := user
   ::login['Pass'] := pass
   ::login['SMTPPass'] := smtp_pass
   ::login['ReplyTo'] := replyTo
   if Empty(smtp_pass)
      ::login['SMTPPass'] := pass
   endif
return nil

method setRecipients(eTo, cc, bcc) class Tsmtp_email
   ::recipients := {'To' => eTo, 'Cc' => AClone(cc), 'Bcc' => AClone(bcc)}
return nil

method prepare(content) class Tsmtp_email
   ::subject := content['assunto']
   ::body := '<html>' + hb_eol() + '<body>' + hb_eol()
   ::body += content['assunto'] + hb_eol() + hb_eol()
   ::body += 'ENVIO DE CT-e' + hb_eol() + hb_eol() + hb_eol()
   ::body += 'Esta empresa não envia SPAM! Este é um e-mail obrigatório por lei.' + hb_eol() + hb_eol()
   ::body += 'Voce esta recebendo um Conhecimento de Transporte Eletrônico de ' + content['nomeRemetente'] + '.' + hb_eol()
   ::body += 'Caso nao queira receber este e-mail, favor entrar em contato pelo e-mail comercial ' + ::recipients['To'] + '.' + hb_eol() + hb_eol()
   ::body += 'O arquivo XML do CT-e encontra-se anexado a este e-mail.' + hb_eol()
   ::body += 'Para verificar a autorização do CT-e junto a SEFAZ, acesse o Portal de consulta através do endereço: https://www.cte.fazenda.gov.br.' + hb_eol() + hb_eol()
   ::body += 'No campo "Chave de acesso", inclua a numeração da chave de acesso abaixo (sem o literal "CTe") e complete a consulta com as informações solicitadas pelo Portal.' + hb_eol() + hb_eol() + hb_eol()
   ::body += 'Chave de acesso:  ' + content["cteChave"] + hb_eol() + hb_eol() + hb_eol() + hb_eol()
   ::body += 'Atenciosamente,' + hb_eol() + hb_eol()
   ::body += content['nomeRemetente'] + hb_eol()
   ::body += content['foneRemetente'] + hb_eol()
   ::body += ::recipients['To'] + hb_eol() + hb_eol()
   ::body += 'TMS Expresso.Cloud' + hb_eol()

   if !Empty(content['portal'])
      ::body += 'Acompanhe sua carga pelo portal' + hb_eol()
      ::body += content['portal'] + hb_eol()
   end

   ::body +=  hb_eol() + hb_eol() + '*** Esse é um e-mail automático. Não é necessário respondê-lo ***' + hb_eol()
   ::body += '</body>' + hb_eol() + '</html>'

return nil

method cc_as_string() class Tsmtp_email
   local eMail := ''
   AEval(::recipients['Cc'], {|mail| eMail += mail + ';'})
return eMail

method bcc_as_string() class Tsmtp_email
   local eMail := ''
   AEval(::recipients['Bcc'], {|mail| eMail += mail + ';'})
return eMail

method attachFile(file) class Tsmtp_email
   AAdd(::attachment, file)
return nil

method is_not_attached() class Tsmtp_email
return (hmg_len(::attachment) == 0)

method sendmail() class Tsmtp_email
   local log
   if ::trace
      log := 'Server: ' + ::server + hb_eol()
      log += 'Port: ' + hb_ntos(::port) + hb_eol()
      log += 'From: ' + ::login['From'] + hb_eol()
      log += 'To: ' + ::recipients['To'] + hb_eol()
      log += 'Cc: ' + array_to_string(::recipients['Cc']) + hb_eol()
      log += 'Bcc: ' + array_to_string(::recipients['Bcc']) + hb_eol()
      log += 'Body: ' + ::msg['Body'] + hb_eol()
      log += 'Subject ' + ::msg['Subject'] + hb_eol()
      log += 'Attachment: ' + array_to_string(::attachment) + hb_eol()
      log += 'User: ' + ::login['User'] + hb_eol()
      log += 'Pass: *********' + hb_eol()
      log += 'POP Server: ' + ::popServer + hb_eol()
      log += 'POP Auth: ' + iif(::popAuth, 'true', 'false') + hb_eol()
      log += 'No Auth: ' + iif(::noAuth, 'true', 'false') + hb_eol()
      log += 'ReplyTo: ' + iif(ValType(::login['ReplyTo']) == "C", ::login['ReplyTo'], '') + hb_eol()
      log += 'TLS: ' + iif(::isTLS, 'true', 'false') + hb_eol()
      log += 'SMTPPass: *******' + hb_eol()
      RegistraLog(log, true)
   endif
return hb_SendMail( ::server,;
                    ::port,;
                    ::login['From'],;
                    ::recipients['To'],;  // string email To
                    ::recipients['Cc'],;  // array emails CC
                    ::recipients['Bcc'],; // array emails BCC
                    ::msg['Body'],;
                    ::msg['Subject'],;
                    ::attachment,;
                    ::login['User'],;
                    ::login['Pass'],;
                    ::popServer,;
                    ::priority,;
                    ::isRead,;
                    ::trace,;
                    ::popAuth,;
                    ::noAuth,;
                    ::timeOut,;
                    ::login['ReplyTo'],;
                    ::isTLS,;
                    ::login['SMTPPass'],;
                    ::charset,;
                    ::encoding)

method reset() class Tsmtp_email
   ::recipients := {'To' => '', 'Cc' => {}, 'Bcc' => {}}
   ::attachment := {}
return nil

method destroy() class Tsmtp_email
   ::recipients := nil
   ::attachment := nil
   ::login := nil
   ::msg := nil
return self

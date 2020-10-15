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
   data from readonly
   data recipients readonly
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
   method setRecipients(to, cc, bcc)
   method setMsg(subject, body)
   method cc_as_string()
   method bcc_as_string()
   method sendmail()
   method reset()
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

method setRecipients(to, cc, bcc) class Tsmtp_email
   ::recipients := {'To' => to, 'Cc' => AClone(cc), 'Bcc' => AClone(bcc)}
return nil

method setMsg(subject, body) class Tsmtp_email
   ::msg := {'Subject' => subject, 'Body' => body}
return nil

method cc_as_string() class Tsmtp_email
   local mail, eMail := ''
   for each mail in ::recipients['Cc']
      eMail += mail + ';'
   next
return eMail

method bcc_as_string() class Tsmtp_email
   local mail, eMail := ''
   for each mail in ::recipients['Bcc']
      eMail += mail + ';'
   next
return eMail

method sendmail() class Tsmtp_email
return hb_SendMail( ::server,;
                    ::port,;
                    ::login['From'],;
                    ::recipients['To'],;
                    ::recipients['Cc'],;
                    ::recipients['Bcc'],;
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
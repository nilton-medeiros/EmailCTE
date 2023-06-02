#include <hmg.ch>
#include "hbclass.ch"
#define true  .T.
#define false .F.

create class TEmailsList

    data emails init {}

    method new(emails_string) constructor
    method is_email() INLINE !(hmg_len(::emails) == 0)
    method getTo()
    method destroy() INLINE ::emails := nil

end class

method new(emails_string) class TEmailsList
    default emails_string := ""
    if !Empty(emails_string)
        emails_string := hb_utf8StrTran(emails_string, ",", ";")
        emails_string := hb_utf8StrTran(emails_string, " ")
        emails_string := hmg_lower(emails_string)
        ::emails := hb_ATokens(emails_string, ";")
    endif
return self

method getTo() class TEmailsList
    local emailTo := ""
    if ::is_email()
        emailTo := ::emails[1]
    endif
return emailTo

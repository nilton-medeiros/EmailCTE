#include <hmg.ch>
#include "fileio.ch"

#define true  .T.
#define false .F.

DECLARE WINDOW Main
DECLARE WINDOW Opcoes

Procedure LoadEmpresas()
          Local n
          Local cSQL, oQuery, nRow

          MsgStatus('Carregando Empresas...', 'dbRefresh' )

          cSQL := "SELECT emp_id, "
          cSQL += "CONCAT(emp_razao_social, '  (', emp_sigla_cia, IF(ISNULL(cid_sigla),'', CONCAT('-',cid_sigla)), ')') AS empresa, "
          cSQL += "emp_nome_fantasia, "
          cSQL += "emp_sigla_cia, "
          cSQL += "emp_cnpj, "
          cSQL += "emp_ambiente_sefaz, "
          cSQL += "emp_fone1, "
          cSQL += "emp_smtp_servidor, "
          cSQL += "emp_smtp_pass, "
          cSQL += "emp_smtp_email, "
          cSQL += "emp_email_contabil, "
          cSQL += "emp_email_comercial, "
          cSQL += "emp_portal, "
          cSQL += "emp_smtp_login, "
          cSQL += "emp_smtp_senha, "
          cSQL += "emp_smtp_porta, "
          cSQL += "emp_smtp_autentica, "
          cSQL += "emp_email_CCO, "
          cSQL += "emp_gmail1_login, "
          cSQL += "emp_gmail1_senha, "
          cSQL += "emp_gmail2_login, "
          cSQL += "emp_gmail2_senha "
          cSQL += "FROM view_empresas "
          cSQL += "WHERE emp_ativa = 1 AND emp_tipo_emitente = 'CTE' AND emp_ambiente_sefaz = 1 "
          cSQL += "ORDER BY empresa"

          oQuery := ExecutaQuery(cSQL)

          if ExecutouQuery( @oQuery, cSQL )
             if oQuery:NetErr()
                RegistraLog('Erro SQL: ' + oQuery:Error() + ' | Descrição do comando SQL: ' + cSQL,, true)
                MsgStatus( 'Erro SQL', 'dbError' )
                PlayExclamation()
                MsgExclamation('Erro SQL: ' + CRLF + oQuery:Error(), 'eMailCTe: Carregando Empresas')
                MsgExclamation({'Descrição do comando SQL: ', CRLF, cSQL, CRLF, 'Ligue para o Suporte ou tente reiniciar o programa'}, 'eMailCTe: Erro de SQL')
                oQuery:Destroy()
                RELEASE WINDOW ALL
             else

                g_oEmpresas:Clean()

                if ( oQuery:LastRec() == 0 )

                   RegistraLog('Empresa não cadastrada no sistema TMS Expresso.Cloud' + hb_eol() + 'SQL: ' + cSQL,, true)
                   MsgStatus('Empresa não cadastrada no sistema TMS Expresso.Cloud','dbError')
                   oQuery:Destroy()
                   MsgExclamation('Empresa não cadastrada no sistema TMS Expresso.Cloud', 'eMailCTe')
                   RELEASE WINDOW ALL

                else

                   FOR n := 1 TO oQuery:LastRec()
                       g_oEmpresas:Adds( oQuery:GetRow(n) )
                   NEXT n

                end
             end

             oQuery:Destroy()

          else
             RegistraLog('Solicitação SQL ao Servidor foi perdida' + CRLF + cSQL,, true)
             MsgStatus( 'Solicitação SQL ao Servidor foi perdida', 'dbError' )
             PlayExclamation()
             MsgExclamation('Solicitação SQL ao Servidor foi perdida', 'eMailCTe: Empresas')
             RELEASE WINDOW ALL
          end

          MsgStatus()

Return

Procedure LoadUsuarios()
          Local n
          Local cSQL, oQuery, oRow

          MsgStatus('Carregando Usuarios...', 'dbRefresh')

          cSQL := "SELECT user_id, user_login, user_senha "
          cSQL += "FROM view_usuarios "
          cSQL += "WHERE user_ativo = 1 AND perm_grupo = 'Administradores' "
          cSQL += "ORDER BY user_login"

          oQuery := ExecutaQuery(cSQL)

          if ExecutouQuery( @oQuery, cSQL )

             if oQuery:NetErr()

                RegistraLog('Erro SQL: ' + oQuery:Error() + ' | Descrição do comando SQL: ' + CRLF + cSQL,, true)
                MsgStatus( 'Erro SQL', 'dbError' )
                PlayExclamation()
                MsgExclamation('Erro SQL: ' + CRLF + oQuery:Error(), 'eMailCTe: Carregando Usuários')
                MsgExclamation({'Descrição do comando SQL: ', CRLF, cSQL, CRLF, 'Ligue para o Suporte ou tente reiniciar o programa'}, 'eMailCTe: Erro de SQL')
                oQuery:Destroy()
                RELEASE WINDOW ALL

             else

                g_aUsuarios := {}

                if oQuery:LastRec() > 0
                   FOR n := 1 TO oQuery:LastRec()
                       oRow := oQuery:GetRow(n)
                       AADD(g_aUsuarios, {'cId' => Str(oRow:FieldGet('user_id'),10), 'login' => AllTrim(oRow:FieldGet('user_login')), 'senha' => AllTrim(oRow:FieldGet('user_senha'))})
                   NEXT
                else
                   RegistraLog('Usuário não cadastrado no sistema TMS Expresso.Cloud')
                   MsgStatus( 'Usuário não cadastrado no sistema TMS Expresso.Cloud', 'dbError' )
                   oQuery:Destroy()
                   MsgExclamation('Usuário não cadastrado no sistema TMS Expresso.Cloud', 'eMailCTe')
                   RELEASE WINDOW ALL
                end

             end

             oQuery:Destroy()

          else

             RegistraLog('Solicitação SQL ao Servidor foi perdida' + CRLF + cSQL,, true)
             MsgStatus( 'Solicitação SQL ao Servidor foi perdida', 'dbError' )
             PlayExclamation()
             MsgExclamation('Solicitação SQL ao Servidor foi perdida', 'eMailCTe: Usuários')
             RELEASE WINDOW ALL

          end

          MsgStatus()

Return

Procedure SysWait( nWait, leMail )
          Local iTime := Seconds()
          Local nLise

          // nWait: segundos

			 DEFAULT nWait  := 2
          DEFAULT leMail := .F.

          nLise := Seconds() - iTime

          Do While nLise < nWait
             inkey(2)
             if ( leMail )
                MsgStatus( "Próximo envio de e-mail em " + LTrim(STR(Int(nWait-nLise))) + "'s", 'emailIcon' )
             end
             DO EVENTS
             nLise := Seconds() - iTime
          EndDo

Return

Procedure MsgStatus( cNotifyTooltip, cNotifyIcon )
         local appName := 'eMailCTe ' + g_cVersao + ' (32-bit) '
          DEFAULT cNotifyTooltip := if( g_oHostConect:Conectado, "Conectado", "Desconectado" )
          DEFAULT cNotifyIcon    := if( g_oHostConect:Conectado, 'dbOn', 'dbOff' )

          if !Empty( g_oEmpresas:sigla_cia)
            Main.NotifyTooltip := appName  + '[' + g_oEmpresas:sigla_cia + ']' + hb_eol() + cNotifyTooltip
          else
            Main.NotifyTooltip := appName + hb_eol() + cNotifyTooltip
          endif
          Main.NotifyIcon := cNotifyIcon

          if IsWIndowActive( Opcoes )
             Opcoes.StatusBar.Item(1) := "Status: " + cNotifyTooltip
             Opcoes.StatusBar.Item(3) := "B.D.: " + if(g_oHostConect:Conectado, 'CONECTADO', 'DESCONECTADO')
             Opcoes.StatusBar.Icon(3) := if(g_oHostConect:Conectado, 'dbOn', 'dbOff')
          end

Return

Function DateToSQL(dDate)
         Local cResult := DTOS(dDate)
         cResult := HB_ULEFT(cResult,4) + '-' + HB_USUBSTR(cResult,5,2) + '-' + HB_USUBSTR(cResult,7,2)
Return (cResult)

Procedure registraLog( cRegistra, lSaltarLinha, lEncrypt )
          Local cLogMsg, nHandle
          Local cDate    := DTOC(Date())
          Local cFolder  := 'log'+hb_ps()
          Local cLogFile := 'email_log_' + HB_URIGHT(cDate,4) + HB_USUBSTR(cDate,4,2) + '.txt'

          DEFAULT lSaltarLinha := false
          DEFAULT lEncrypt := false

          cRegistra := HB_UTF8STRTRAN( cRegistra, CRLF, CRLF + Space(20) )

          if lEncrypt
            cRegistra := CharXor(cRegistra, "MyKeySendMail")
          endif

			 if hb_FileExists( cFolder+cLogFile )
             nHandle := fOpen( cFolder+cLogFile, FO_WRITE )
             fSeek( nHandle, 0, FS_END )
          else
             nHandle := fCreate( cFolder+cLogFile, FC_NORMAL )
             fWrite( nHandle, 'eMailCTe Log do sistema "TMS.Client Expresso.Cloud" - v.' + g_cVersao + hb_eol() + hb_eol() )
          end

          cLogMsg := cDate + ' ' + time() + HMG_PADR( ' | Processo: ' + AllTrim(ProcName(1)), 35 ) + '| Linha: ' + Transform(ProcLine(1), "9999")
          cLogMsg += ' | ' + cRegistra + hb_eol()

			 if is_true( lSaltarLinha )
			    cLogMsg := CRLF + cLogMsg + CRLF
			 end

          fWrite(nHandle, cLogMsg)
          fClose(nHandle)

Return

Function IP_Externo()
         //System.Clipboard := html + url + vRet      // Debug
RETURN PegaIP_ex( ReadPage_ler( 'http://www.meuip.com.br/' ) )

FUNCTION PegaIP_ex(cHtml)
         LOCAL Pos, PosF

         Pos := At('INNERHTML = "', Upper(cHtml) )

         IF Pos < 1
           RETURN ''
         ENDIF

         Pos   += 11
         cHtml := HB_USUBSTR( cHtml, Pos )
         PosF  := At('"',cHtml) - 1
         cHtml := HB_ULEFT(cHtml,PosF)

RETURN cHtml

FUNCTION ReadPage_ler( cUrl )
         LOCAL oUrl, oCli, cRes := ''

         BEGIN SEQUENCE

           oUrl := TUrl():New( cUrl )

           IF EMPTY( oUrl )
              //MsgDebug('Não abriu url, linha 101', cUrl, oUrl)
              BREAK
           ENDIF

           oCli := TIpClientHttp():New( oUrl )
           //oCli = TIPClient():New( oUrl ) // para uso em xharbour 9970

           IF EMPTY( oCli )
              //MsgDebug('Não abriu objeto oCli, linha 109', oCli)
              BREAK
           ENDIF

           oCli:nConnTimeout = 20000

           IF !oCli:Open( oUrl )
              BREAK
           ENDIF

           cRes := oCli:Read()
           oCli:Close()
         END SEQUENCE

RETURN cRes

Function mysql_escape(_string)
         _string := AllTrim(_string)
         _string := mysql_escape_string(_string)
Return TirarAcentos(_string)

Function TirarAcentos(cTexto)

         cTexto := HB_UTF8STRTRAN(cTexto, "Ã", "A")
         cTexto := HB_UTF8STRTRAN(cTexto, "Á", "A")
         cTexto := HB_UTF8STRTRAN(cTexto, "À", "A")
         cTexto := HB_UTF8STRTRAN(cTexto, "Â", "A")
         cTexto := HB_UTF8STRTRAN(cTexto, "ã", "a")
         cTexto := HB_UTF8STRTRAN(cTexto, "á", "a")
         cTexto := HB_UTF8STRTRAN(cTexto, "à", "a")
         cTexto := HB_UTF8STRTRAN(cTexto, "â", "a")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ä", "A")
         cTexto := HB_UTF8STRTRAN(cTexto, "ä", "a")
         cTexto := HB_UTF8STRTRAN(cTexto, "É", "A")
         cTexto := HB_UTF8STRTRAN(cTexto, "È", "E")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ê", "E")
         cTexto := HB_UTF8STRTRAN(cTexto, "é", "E")
         cTexto := HB_UTF8STRTRAN(cTexto, "è", "e")
         cTexto := HB_UTF8STRTRAN(cTexto, "ê", "e")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ë", "E")
         cTexto := HB_UTF8STRTRAN(cTexto, "ë", "e")
         cTexto := HB_UTF8STRTRAN(cTexto, "Í", "I")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ì", "I")
         cTexto := HB_UTF8STRTRAN(cTexto, "Î", "I")
         cTexto := HB_UTF8STRTRAN(cTexto, "í", "i")
         cTexto := HB_UTF8STRTRAN(cTexto, "ì", "i")
         cTexto := HB_UTF8STRTRAN(cTexto, "î", "i")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ï", "I")
         cTexto := HB_UTF8STRTRAN(cTexto, "ï", "i")
         cTexto := HB_UTF8STRTRAN(cTexto, "Õ", "o")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ó", "o")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ò", "o")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ô", "o")
         cTexto := HB_UTF8STRTRAN(cTexto, "õ", "o")
         cTexto := HB_UTF8STRTRAN(cTexto, "ó", "o")
         cTexto := HB_UTF8STRTRAN(cTexto, "ò", "o")
         cTexto := HB_UTF8STRTRAN(cTexto, "ô", "o")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ö", "O")
         cTexto := HB_UTF8STRTRAN(cTexto, "ö", "o")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ú", "U")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ù", "U")
         cTexto := HB_UTF8STRTRAN(cTexto, "Û", "U")
         cTexto := HB_UTF8STRTRAN(cTexto, "ú", "U")
         cTexto := HB_UTF8STRTRAN(cTexto, "ù", "U")
         cTexto := HB_UTF8STRTRAN(cTexto, "û", "U")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ü", "U")
         cTexto := HB_UTF8STRTRAN(cTexto, "ü", "u")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ý", "Y")
         cTexto := HB_UTF8STRTRAN(cTexto, "ý", "y")
         cTexto := HB_UTF8STRTRAN(cTexto, "ÿ", "y")
         cTexto := HB_UTF8STRTRAN(cTexto, "Ç", "C")
         cTexto := HB_UTF8STRTRAN(cTexto, "ç", "c")
         cTexto := HB_UTF8STRTRAN(cTexto, "º", "o.")
         cTexto := HB_UTF8STRTRAN(cTexto, "", "A")
         cTexto := HB_UTF8STRTRAN(cTexto, "", "a")
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(195) + Chr(173), "i" ) // i acentuado minusculo
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(195) + Chr(135), "C" ) // c cedilha maiusculo
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(195) + Chr(141), "I" ) // i acentuado maiusculo
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(195) + Chr(163), "a" ) // a acentuado minusculo
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(195) + Chr(167), "c" ) // c acentuado minusculo
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(195) + Chr(161), "a" ) // a acentuado minusculo
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(195) + Chr(131), "A" ) // a acentuado maiusculo
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(194) + Chr(186), "o." ) // numero simbolo
         // so pra corrigir no MySql
         cTexto := HB_UTF8STRTRAN( cTexto, "+" + Chr(129), "A" )
         cTexto := HB_UTF8STRTRAN( cTexto, "+" + Chr(137), "E" )
         cTexto := HB_UTF8STRTRAN( cTexto, "+" + Chr(131), "A" )
         cTexto := HB_UTF8STRTRAN( cTexto, "+" + Chr(135), "C" )
         cTexto := HB_UTF8STRTRAN( cTexto, "?" + Chr(167), "c" )
         cTexto := HB_UTF8STRTRAN( cTexto, "?" + Chr(163), "a" )
         cTexto := HB_UTF8STRTRAN( cTexto, "?" + Chr(173), "i" )
         cTexto := HB_UTF8STRTRAN( cTexto, "?" + Chr(131), "A" )
         cTexto := HB_UTF8STRTRAN( cTexto, "?" + Chr(161), "a" )
         cTexto := HB_UTF8STRTRAN( cTexto, "?" + Chr(141), "I" )
         cTexto := HB_UTF8STRTRAN( cTexto, "?" + Chr(135), "C" )
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(195) + Chr(156), "a" )
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(195) + Chr(159), "A" )
         cTexto := HB_UTF8STRTRAN( cTexto, "?" + Chr(129), "A" )
         cTexto := HB_UTF8STRTRAN( cTexto, "?" + Chr(137), "E" )
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(195) + "?", "C" )
         cTexto := HB_UTF8STRTRAN( cTexto, "?" + Chr(149), "O" )
         cTexto := HB_UTF8STRTRAN( cTexto, "?" + Chr(154), "U" )
         cTexto := HB_UTF8STRTRAN( cTexto, "+" + Chr(170), "o" )
         cTexto := HB_UTF8STRTRAN( cTexto, "?" + Chr(128), "A" )
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(195) + Chr(166), "e" )
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(135) + Chr(227), "ca" )
         cTexto := HB_UTF8STRTRAN( cTexto, "n" + Chr(227), "na" )
         cTexto := HB_UTF8STRTRAN( cTexto, Chr(162), "o" )

Return (cTexto)

Function DelItemArray( aVelho, nItem )
         Local k, aTemp := {}

         FOR k := 1 TO HMG_LEN( aVelho )
             if !( k == nItem )
                AAdd( aTemp, aVelho[k] )
             end
         NEXT k

Return AClone( aTemp )

Function is_true( xBoolean )
   local result := false

   switch ValType(xBoolean)
      case 'L'
         result := xBoolean
         exit
      case 'N'
         result := (xBoolean > 0)
         exit
      case 'C'
         result := !Empty(xBoolean)
   endswitch

return result

Function array_to_string(arrayOfString)
   local result := '', conteudo
   if ValType(arrayOfString) == "A" .and. Len(arrayOfString) > 0
      for each conteudo in arrayOfString
         switch ValType(conteudo)
            case "C"
               exit
            case "D"
               conteudo := DToC(conteudo)
               exit
            case "N"
               conteudo := hb_ntos(conteudo)
               exit
            case "L"
               conteudo := iif(conteudo, 'true', 'false')
               exit
            otherwise
               conteudo := 'Tipo do conteúdo do array não é string. Tipo: "' + ValType(conteudo) + '"'
         endswitch
         result += ' | ' + conteudo
      next
   elseif ValType(arrayOfString) == "C"
      result := arrayOfString
   endif
return LTrim(result)

function ifNull(ver, _default)
   if ValType(ver) == "U" .or. Empty(ver)
      return default
   endif
return ver
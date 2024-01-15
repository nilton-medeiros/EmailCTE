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
                RegistraLog('Erro SQL: ' + oQuery:Error() + ' | Descrição do comando SQL: ' + cSQL)
                MsgStatus( 'Erro SQL', 'dbError' )
                PlayExclamation()
                MsgExclamation('Erro SQL: ' + CRLF + oQuery:Error(), 'eMailCTe: Carregando Empresas')
                MsgExclamation({'Descrição do comando SQL: ', CRLF, cSQL, CRLF, 'Ligue para o Suporte ou tente reiniciar o programa'}, 'eMailCTe: Erro de SQL')
                oQuery:Destroy()
                RELEASE WINDOW ALL
             else

                g_oEmpresas:Clean()

                if ( oQuery:LastRec() == 0 )

                   RegistraLog('Empresa não cadastrada no sistema TMS Expresso.Cloud' + hb_eol() + 'SQL: ' + cSQL)
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
             RegistraLog('Solicitação SQL ao Servidor foi perdida' + CRLF + cSQL)
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

                RegistraLog('Erro SQL: ' + oQuery:Error() + ' | Descrição do comando SQL: ' + CRLF + cSQL)
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

             RegistraLog('Solicitação SQL ao Servidor foi perdida' + CRLF + cSQL)
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
            cRegistra := "[ENCRYPT START =>]" + hb_eol() + auxEncrypt(cRegistra) + hb_eol() + "[<= ENCRYPT END]"
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

Function array_to_string(arrayOfString)
   local conteudo
   local result := ''

   if ValType(arrayOfString) == "A" .and. hmg_len(arrayOfString) > 0

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
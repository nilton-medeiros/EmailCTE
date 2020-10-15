
/*

 *****************************************************************************************
 * Função personalizada para  envio  de  e-Mail  usando
 * Componente CDOsys da Microsoft em servidores Windows
 *       2016/09/21  -  by Nilton G. Medeiros
 *
 * Nota: ENVIO DA MENSAGEM NO FORMATO TEXTO (e-mail simples)
 *       PARA ENVIO DA MENSAGEM NO FORMATO HTML, ALTERE O TextBody PARA HtmlBody
 *       e procure informações sobre o componente CDOsys Microsoft e TAG's HTMLs
 *
 * fonte: http://search.msdn.microsoft.com/search/default.aspx?siteId=0&tab=0&query=cdosys
 *
 *****************************************************************************************

 */

#include <hmg.ch>
#include <hbcompat.ch>

Function  My_SendMail( hSmtpDados )
          Local i, cDebug, cAnexo
          Local lRet := .F.
          Local cMsg
          Local oCfg := win_OleCreateObject( "CDO.Configuration" )
          Local oErroMail, oMsg

          //--> CONFIGURAÇOES DE E-MAIL

          TRY

             WITH OBJECT oCfg:Fields
               :Item("http://schemas.microsoft.com/cdo/configuration/smtpserver"):Value       := hSmtpDados['Servidor']
               :Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport"):Value   := hSmtpDados['Porta']
               :Item("http://schemas.microsoft.com/cdo/configuration/sendusing"):Value        := 2      // sendusing : cdoSendUsingPort, valor 2, para enviar a mensagem usando a rede
               :Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"):Value := hSmtpDados['Autentica']
               :Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl"):Value       := hSmtpDados['SSL']
               :Item("http://schemas.microsoft.com/cdo/configuration/sendusername"):Value     := hSmtpDados['Login']
               :Item("http://schemas.microsoft.com/cdo/configuration/sendpassword"):Value     := hSmtpDados['Senha']
               :Update()
             END WITH

             cMsg := 'CDO: Objeto de eMail Configurado com sucesso:'
             cMsg += ' | De:        ' + hSmtpDados['De']
             cMsg += ' | Para:      ' + hSmtpDados['Para']

             if !Empty( hSmtpDados['CC'] )
                cMsg += ' | CC:        ' + hSmtpDados['CC']
             end
             if !Empty( hSmtpDados['BCC'] )
                cMsg += ' | BCC:       ' + hSmtpDados['BCC']
             end

             cMsg += ' | Responder: ' + hSmtpDados['Responder']
             cMsg += ' | Assunto:   ' + hSmtpDados['Assunto']

             RegistraLog( cMsg )

/*
 *           RegistraLog( 'CDO: Objeto de eMail Configurado com sucesso: ' + CRLF +;
 *                        '| Servidor:  ' + hSmtpDados['Servidor'] + CRLF + ;
 *                        '| Porta:     ' + LTrim(STR(hSmtpDados['Porta'])) + ;
 *                        '| Autentica: ' + if(hSmtpDados['Autentica'], 'SIM', 'NÃO') + ' | SSL: ' + if(hSmtpDados['SSL'], 'SIM', 'NÃO') + CRLF + ;
 *                        '| Login:     ' + hSmtpDados['Login'] + CRLF + ;
 *                        '| De:        ' + hSmtpDados['De'] + CRLF +  ;
 *                        '| Para:      ' + hSmtpDados['Para'] + CRLF +  ;
 *                        '| CC:        ' + hSmtpDados['CC'] + CRLF +  ;
 *                        '| BCC:       ' + hSmtpDados['BCC'] + CRLF +  ;
 *                        '| Responder: ' + hSmtpDados['Responder'] + CRLF + ;
 *                        '| Assunto:   ' + hSmtpDados['Assunto'] ;
 *                      )
*/
             lRet := .T.

          CATCH oErroMail

             cDebug := 'ENVIO DE e-MAILs: Erro ao configurar e-mail!' + CRLF + ;
                       'Error:     ' + Transform(oErroMail:GenCode  , nil) + CRLF + ;
                       'SubC:      ' + Transform(oErroMail:SubCode  , nil) + CRLF + ;
                       'OSCode:    ' + Transform(oErroMail:OsCode   , nil) + CRLF + ;
                       'SubSystem: ' + Transform(oErroMail:SubSystem, nil) + CRLF + ;
                       'Mensagem:  ' + oErroMail:Description + CRLF + CRLF + ;
                       'Server:    ' + hSmtpDados['Servidor'] + CRLF + ;
                       'Porta:     ' + LTrim(Str(hSmtpDados['Porta'])) + CRLF + ;
                       'Authentic: ' + if(hSmtpDados['Autentica'], 'SIM', 'NÃO') + CRLF + ;
                       'SSL:       ' + if(hSmtpDados['SSL'], 'SIM', 'NÃO') + CRLF + ;
                       'Login:     ' + hSmtpDados['Login'] + CRLF + ;
                       'Senha:     ' + Left(hSmtpDados['Senha'],2) + '****' + CRLF + ;
                       'From:      ' + hSmtpDados['De'] + CRLF + ;
                       'To:        ' + hSmtpDados['Para'] + CRLF + ;
                       'CC:        ' + hSmtpDados['CC'] + CRLF + ;
                       'BCC:       ' + hSmtpDados['BCC'] + CRLF + ;
                       'Responder: ' + hSmtpDados['Responder'] + CRLF + ;
                       'Subject:   ' + hSmtpDados['Assunto'] + CRLF + CRLF

             cDebug += 'ANEXO(s)' + CRLF
             cDebug += '--------' + CRLF

             FOR EACH cAnexo IN hSmtpDados['Anexos']
                 cDebug += cAnexo + CRLF
             NEXT

             RegistraLog( cDebug )

/*
             if (lAlerta)
                MsgExclamation({'Houve erro ao configurar o e-mail.', CRLF, CRLF  ,       ;
                                'Error:     ', Transform(oErroMail:GenCode  , nil), CRLF, ;
                                'SubC:      ', Transform(oErroMail:SubCode  , nil), CRLF, ;
                                'OSCode:    ', Transform(oErroMail:OsCode   , nil), CRLF, ;
                                'SubSystem: ', Transform(oErroMail:SubSystem, nil), CRLF, ;
                                'Mensagem:  ', oErroMail:Description}, 'Erro configurando e-mail!')
             end
*/
          END

          //--> FIM DAS CONFIGURAÇOES.

          if ( lRet )

             // --> ENVIA E-MAIL

             TRY

               oMsg := win_OleCreateObject( "CDO.Message" )

               WITH OBJECT oMsg

                  :Configuration := oCfg
                  :From          := hSmtpDados['De']
                  :To            := hSmtpDados['Para']
                  :CC            := hSmtpDados['CC']
                  :BCC           := hSmtpDados['BCC']
                  :ReplyTo       := hSmtpDados['Responder']
                  :Subject       := hSmtpDados['Assunto']

/*
                  * ------------------------------------------------------------
                  * Aqui adiciona a imagem ao corpo da mensagem qdo formato HTML
                  * ------------------------------------------------------------
                  IF !Empty(cImagem)
                   :AddRelatedBodyPart(hb_DirBase()+"img"+hb_PS()+cImagem, cImagem, 1)
                   :Fields:Item("urn:schemas:mailheader:Content-ID"):Value := "<"+cImagem+">"
                   :Fields:Item("urn:schemas:mailheader:Content-Disposition"):Value := "inline"
                   :Fields:Update()
                  ENDIF

                  :HTMLBody := cMsg // + QuebraHTML + IF(!Empty(cImagem), cImagem1, "")
                  * ------------------------------------------------------------------------------------------------
*/

                  :TEXTBody := hSmtpDados['Mensagem']

                  FOR EACH cAnexo IN hSmtpDados['Anexos']
                      /* A Classe CDOsys para essa propriedade, tem que por o caminho completo do arquivo
                       * Função nativa do harbour para trazer o caminho completo desde a letra do drive até a pasta corrente: "hb_cwd()" */
                      :AddAttachment(cAnexo)
                  NEXT

                  :Fields("urn:schemas:mailheader:disposition-notification-to"):Value := hSmtpDados['De']
                  :Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"):Value := hSmtpDados['Autentica']
                  :Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl"):Value := hSmtpDados['SSL']
                  :Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver"):Value := hSmtpDados['Servidor']
                  :Fields:update()
                  :Send()

               END WITH

               //MsgInfo( 'E-mail enviado com sucesso.', 'Sucesso!')

             CATCH oErroMail

               cDebug := 'ENVIO DE e-MAILs: Nao foi possivel enviar mensagem!' + CRLF
               cDebug += 'Erro:      ' + TirarAcentos(oErroMail:Description) + CRLF + CRLF
               cDebug += 'Server:    ' + hSmtpDados['Servidor'] + CRLF
               cDebug += 'Porta:     ' + LTrim(Str(hSmtpDados['Porta'])) + CRLF
               cDebug += 'Authentic: ' + if(hSmtpDados['Autentica'], 'SIM', 'NAO') + CRLF
               cDebug += 'SSL:       ' + if(hSmtpDados['SSL'], 'SIM', 'NAO') + CRLF
               cDebug += 'Login:     ' + hSmtpDados['Login'] + CRLF
               cDebug += 'Senha:     ' + Left(hSmtpDados['Senha'],2) + '****' + CRLF
               cDebug += 'From:      ' + hSmtpDados['De'] + CRLF
               cDebug += 'To:        ' + hSmtpDados['Para'] + CRLF
               cDebug += 'CC:        ' + hSmtpDados['CC'] + CRLF
               cDebug += 'BCC:       ' + hSmtpDados['BCC'] + CRLF
               cDebug += 'Responder: ' + hSmtpDados['Responder'] + CRLF
               cDebug += 'Subject:   ' + hSmtpDados['Assunto'] + CRLF + CRLF
               cDebug += 'ANEXO(s)' + CRLF
               cDebug += '--------' + CRLF

               FOR EACH cAnexo IN hSmtpDados['Anexos']
                   cDebug += cAnexo + CRLF
               NEXT

               RegistraLog( cDebug )

/*
               MsgExclamation( {'Não foi possível enviar a mensagem: ' + hSmtpDados['Assunto'], CRLF, ;
                                'para o email: ', hSmtpDados['Para'], '.', CRLF, ;
                                'Erro: ', oErroMail:Description}, 'Erro!' )
*/
               lRet := .F.

             END

          end

Return ( lRet )
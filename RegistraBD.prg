#include <hmg.ch>

declare window RegistraBD

procedure FormRegistraBD()
   LOAD WINDOW RegistraBD
      ON KEY ESCAPE OF RegistraBD ACTION registraBD_Form_onEscape()
   RegistraBD.Center
   RegistraBD.Activate
return

Procedure registraBD_Button_Salvar_Action()
          Local aIP := RegistraBD.IpAddress_host.Value

          if !( aIP[1] == 0 ) .and. !Empty( RegistraBD.Text_DataBase.Value ) .and. !Empty( RegistraBD.Text_User.Value ) .and. !Empty( RegistraBD.Text_Password.Value )

             g_oHostConect:Name     := IPArrayToString( aIP )
             g_oHostConect:DataBase := AllTrim( RegistraBD.Text_DataBase.Value )
             g_oHostConect:User     := AllTrim( RegistraBD.Text_User.Value )
             g_oHostConect:Password := AllTrim( RegistraBD.Text_Password.Value )

             RegistryWrite( g_cRegPath + "Host\Name"    , g_oHostConect:Name     )
             RegistryWrite( g_cRegPath + "Host\DataBase", g_oHostConect:DataBase )
             RegistryWrite( g_cRegPath + "Host\User"    , g_oHostConect:User     )
             RegistryWrite( g_cRegPath + "Host\Password", CharXor( g_oHostConect:Password, 'tms2017' ) )

             if ConectaMySQL()
                RegistraBD.Release
             else
                MsgExclamation( {'Acesso negado para usuário ', g_oHostConect:User,'.', CRLF+CRLF, 'Verifique as informações digitadas ou chame o suporte.' }, 'Informações incorretas!' )
             end

          else
             MsgExclamation( 'Preencher todos os campos', 'Dados insuficientes!' )
          end

Return

Function IPArrayToString( aIP )
         Local cIP := hb_ntos(aIP[1])
         cIP += '.' + hb_ntos(aIP[2])
         cIP += '.' + hb_ntos(aIP[3])
         cIP += '.' + hb_ntos(aIP[4])
Return ( cIP )

procedure registraBD_Form_onEscape()
   RegistraBD.Release
return
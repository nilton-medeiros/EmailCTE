#include <hmg.ch>
#include <common.ch>
#include <i_MsgBox.ch>

Procedure Main

          // Variáveis Globais (Públicas)

          Public g_oServer      := NIL
          Public g_oHostConect  := HostConect():New()
          Public g_oEmpresas    := InfoEmpresa():New()

          Public g_aMaiLogEvent := {}
          Public g_aMaiComErros := {}
          Public g_aUsuarios    := {}
          Public g_hEMailServer := hb_Hash( 'dia', date(), 'server_1', 0, 'server_2', 0 )

          Public g_iTimer       := Seconds()
          Public g_iTimer_Erro  := Seconds()

          Public g_cRegPath     := "HKEY_CURRENT_USER\SOFTWARE\Sistrom\eMailCTE\"
          Public g_cVersao      := "4.5.8"

          // Em algumas rotinas como download ou insert no BD, não poderá aboartar o sistema por causa da atualização
          Public g_lStopExecution := .T.

          IF HMG SUPPORT UNICODE RUN

          SET MULTIPLE OFF
          SET CODEPAGE TO UNICODE
          SET LANGUAGE TO PORTUGUESE
          SET TOOLTIPSTYLE BALLOON
          SET CENTURY ON
          SET EPOCH TO 1990
          SET DATE BRITISH
          SET NAVIGATION EXTENDED

          REQUEST DBFCDX

          dbSetDriver( "DBFCDX" )

          if dbSetDriver() <> "DBFCDX"
             RegistraLog( "Driver [DBFCDX] nao foi instalado!" )
             //MsgStop( "Driver [DBFCDX] nao foi instalado!", "Erro DBFCDX" )
             QUIT
          end

          if !hb_DirExists( '..\Log_Erros\' )
             DirMake( '..\Log_Erros\' )
          end

          if !hb_DirExists( '..\inter_systems\' )
             DirMake( '..\inter_systems\' )
          end

          if !hb_FileExists( '..\inter_systems\intersys.DBF' )
             dbCreate( '..\inter_systems\intersys', {{"APLICATIVO","C",50,0},{"STOP_EXEC","L",1,0},{"MENSAGEM","C",100,0},{"DIA","N",2,0},{"HORA","C",5,0}} )
          end

          dbCloseAll()

          if !hb_FileExists( '..\inter_systems\intersys.CDX' )
             // Nova area, Exclusivo, e não ReadOnly
             dbUseArea( .T., 'DBFCDX', '..\inter_systems\intersys', 'InterSys', .F., .F. )
             INDEX ON InterSys->APLICATIVO TAG TgAplicativo TO '..\inter_systems\intersys.CDX' UNIQUE
             dbCloseAll()
          end

          UseInterSys()

          if ( RegistryRead( g_cRegPath + "NameExec" ) == NIL )       ; RegistryWrite( g_cRegPath + "NameExec", "eMailCTe.exe" ); end
          if ( RegistryRead( g_cRegPath + "Path" ) == NIL )           ; RegistryWrite( g_cRegPath + "Path", hb_cwd() )          ; end
          if ( RegistryRead( g_cRegPath + "Version" ) == NIL )        ; RegistryWrite( g_cRegPath + "Version", g_cVersao )      ; end
          if ( RegistryRead( g_cRegPath + "Auto_Execution" ) == NIL ) ; RegistryWrite( g_cRegPath + "Auto_Execution", 1 )       ; end
          if ( RegistryRead( g_cRegPath + "Em_Execucao" ) == NIL )    ; RegistryWrite( g_cRegPath + "Em_Execucao", 0 )          ; end
          if ( RegistryRead( g_cRegPath + "Cliente" ) == NIL ); RegistryWrite( g_cRegPath + "Cliente", 'Definir Cliente' ) ; end
          if ( RegistryRead( g_cRegPath + "Monitoring\Stop_Execution" ) == NIL ) ; RegistryWrite( g_cRegPath + "Monitoring\Stop_Execution", 0 )       ; end
          if ( RegistryRead( g_cRegPath + "Monitoring\das" ) == NIL )  ; RegistryWrite( g_cRegPath + "Monitoring\das", "22:00" )  ; end
          if ( RegistryRead( g_cRegPath + "Monitoring\ate" ) == NIL )  ; RegistryWrite( g_cRegPath + "Monitoring\ate", "08:00" )  ; end
          if ( RegistryRead( g_cRegPath + "Monitoring\frequencia" ) == NIL ); RegistryWrite( g_cRegPath + "Monitoring\frequencia", 1 ) ; end
/*
          MsgInfo(OS() + CRLF + ;
          Hb_Compiler() + CRLF + ;
          Version() + CRLF + ;
          MiniGuiVersion())
*/
          if is_true( RegistryRead( g_cRegPath + "Monitoring\Stop_Execution" ) )
             MessageBoxTimeout('O programa NÃO será executado. Ambiente de teste ou em atualização!', 'eMailCTe TMS Expresso.Cloud (' + g_cVersao + ')', MB_ICONEXCLAMATION, 10000 )
             QUIT
          end

          if is_true( RegistryRead( g_cRegPath + 'Em_Execucao') )
             RegistraLog('eMailCTE ' + g_cVersao + ': Sistema foi interrompido anteriormente')
          else
             RegistryWrite( g_cRegPath + 'Em_Execucao', 1 )
          end

          RegistraLog('eMailCTE ' + g_cVersao + ': Sistema iniciado (em execução)' )

          LOAD WINDOW Main
          CENTER WINDOW Main
          ACTIVATE WINDOW Main

          // Depois do Activate nada é mais executado aqui nesta Procedure

Return

Procedure sobre()
          ShellAbout( "eMailCTe", ;
                      "eMailCTe - TMS Expresso.Cloud: versão " + ;
                      g_cVersao + CRLF + Chr(169) + ;
                      " by Sistrom Sistemas Web, 2016", ;
                    LoadTrayIcon(GetInstance(), "MAIN") )
Return

Procedure main_form_oninit()
          Local cFolderLog := 'log'+hb_ps()
          Local cFantasia, cPasta
          Local i

          SetProperty( "Main", "Timer_Mail", "Enabled", .F. )
          SetProperty( "Main", "Timer_yes_update", "Enabled", .F. )
          if hb_DirExists(cFolderLog)
             AEVAL( DIRECTORY( "log\*.*"), { |aFile| IF( aFile[3]<=Date()-60, FERASE( "log\"+aFile[1] ), NIL ) } )
          else
             DirMake(cFolderLog)
          end

          if ( RegistryRead( g_cRegPath + "Host\Name" ) == NIL ) .or. ;
             ( RegistryRead( g_cRegPath + "Host\DataBase" ) == NIL ) .or. ;
             ( RegistryRead( g_cRegPath + "Host\User" ) == NIL ) .or. ;
             ( RegistryRead( g_cRegPath + "Host\Password" ) == NIL )

             FormRegistraBD()

             if Empty( g_oHostConect:Name ) .or. Empty( g_oHostConect:User ) .or. Empty( g_oHostConect:DataBase ) .or. Empty( g_oHostConect:Password )
                MsgStop( {'Dados insuficientes para conexão com Host:', CRLF+CRLF, 'Servidor: ', g_oHostConect:Name, CRLF, 'Banco Dados: ', g_oHostConect:DataBase }, 'Conexão com Host!' )
                RELEASE WINDOW ALL
             end

          else

             g_oHostConect:Name     := RegistryRead( g_cRegPath + "Host\Name" )
             g_oHostConect:DataBase := RegistryRead( g_cRegPath + "Host\DataBase" )
             g_oHostConect:User     := RegistryRead( g_cRegPath + "Host\User" )
             g_oHostConect:Password := CharXor( RegistryRead( g_cRegPath + "Host\Password" ), 'tms2017' )

             if ( RegistryRead( g_cRegPath + "Version" ) == NIL ) .or. !( RegistryRead( g_cRegPath + "Version" ) == g_cVersao )
                RegistryWrite( g_cRegPath + "Version", g_cVersao )
             end

          end

          if ConectaMySQL()

             LoadEmpresas()
             LoadUsuarios()
             MessageBoxTimeout('O eMailCTe ficará oculto na barra de tarefas.', 'eMailCTe TMS Expresso.Cloud ' + g_cVersao + ': Aviso!', MB_ICONEXCLAMATION, 7000 )

          else
             MsgStop('Sem conexão com Banco de Dados!','eMailCTe: Falha ao Conectar')
             RegistraLog('Falha ao Conectar. Sem conexão com Banco de Dados!')
             RELEASE WINDOW ALL
          end

          if !hb_DirExists( '..\ctes_baixados\' )
             DirMake( '..\ctes_baixados\' )
          end

          FOR i := 1 TO g_oEmpresas:QtdeEmpresas
              // Configuração das empresas e criação das pastas pdf & xml
              g_oEmpresas:SetFolders(i)
          NEXT i

          SetProperty( "Main", "Timer_Mail", "Enabled", .T. )
          SetProperty( "Main", "Timer_yes_update", "Enabled", .T. )

Return

Procedure main_timer_Mail_action
          // 1200000  == 20 minutos
          //   60000  ==  60 segundos = 1 minuto
          SetProperty( "Main", "Timer_Mail", "Enabled", .F. )
          MonitoraMails()
          MonitoraErros()
          SetProperty( "Main", "Timer_Mail", "Enabled", .T. )
          DO EVENTS
Return

Procedure main_timer_yes_update_action
          // 10000   == 10 segundos
          if ( g_lStopExecution )
            if is_true(RegistryRead(g_cRegPath + "Monitoring\Stop_Execution"))
               desligar()
            endif
             SetProperty( "Main", "Timer_yes_update", "Enabled", .F. )
             UseInterSys()
             SetProperty( "Main", "Timer_yes_update", "Enabled", .T. )
          end
Return

Procedure desligar()
          SetProperty( "Main", "Timer_Mail", "Enabled", .F. )
          SetProperty( "Main", "Timer_yes_update", "Enabled", .F. )
          if !( g_lStopExecution )
                g_lStopExecution := .T.
          else
             RegistraLog('Sistema fechado pelo usuário' )
             RELEASE WINDOW ALL
          end
          SetProperty( "Main", "Timer_Mail", "Enabled", .T. )
          SetProperty( "Main", "Timer_yes_update", "Enabled", .T. )
Return

Procedure main_form_onrelease()
          RegistryWrite( g_cRegPath + 'Em_Execucao', 0 )
          DesconectaMySQL()
          dbCloseAll()
Return

Procedure UseInterSys()
          Local nHI, nMI, nHF, nMF

          dbCloseAll()

          // Nova area, Compartilhado, e não ReadOnly
          dbUseArea( .T., 'DBFCDX', '..\inter_systems\intersys', 'InterSys', .T., .F. )

          CheckError( 'intersys.dbf')

          SET INDEX TO ..\inter_systems\intersys

          CheckError( 'de índice intersys.CDX' )

          dbSelectArea( 'InterSys' )
          dbSetOrder( 'TgAplicativo' )

          if InterSys->( dbSeek('EMAILCTE.EXE') )

             if ( InterSys->STOP_EXEC )

                if ( InterSys->DIA == Day( Date() ) )

                   nHI := Val( Left( InterSys->HORA, 2 ) )
                   nMI := Val( Right( InterSys->HORA, 2 ) )
                   nHF := Val( Left( Time(), 5 ) )
                   nMF := Val( Substr( Time(), 4, 2 ) )

                   if ( ((nHF-nHI) * 60) + (nMF-nMI) ) < 20
                      // Parada solicitada em menos de 20 minutos, parar sistema
                      RegistraLog( 'Parada forçada do sistema para atualização remota!' )
                      RELEASE WINDOW ALL
                   end

                end

                // Parada solicitada expirada a mais de 20 minutos, ignorar parada.
                InterSys->( RLock() )
                InterSys->STOP_EXEC := false
                InterSys->MENSAGEM  := 'Update expirado'

             end

          else

             InterSys->( dbAppend() )
             InterSys->APLICATIVO := 'EMAILCTE.EXE'
             InterSys->STOP_EXEC  := false
             InterSys->MENSAGEM   := 'Processo de inicializacao'

          end

          InterSys->( dbUnLockAll() )
          InterSys->( dbCommitAll() )

          dbCloseAll()

Return

Procedure CheckError( cMsg )
          Local lNetErro := NetErr()

          lNetErro := IF( ValType( lNetErro ) == "L", lNetErro, .T. )

          if ( lNetErro )
             MsgStop( 'Erro de rede abrindo arquivo ' + cMsg + CRLF + 'Código do erro: ' + LTrim(Str(DosError())), 'Erro de Rede' )
             RegistraLog( 'Erro abrindo arquivo ' + cMsg )
             RELEASE WINDOW ALL
          elseif ( DosError() == 5 )
             MsgStop( 'Acesso negado abrindo arquivo ' + cMsg, 'Erro de Rede' )
             RegistraLog( 'Acesso negado abrindo arquivo ' + cMsg )
             RELEASE WINDOW ALL
          elseif ( DosError() == 32 )
             MsgStop( 'Acesso de compartilhamento abrindo arquivo ' + cMsg, 'Erro de Rede' )
             RegistraLog( 'Acesso negado abrindo arquivo ' + cMsg )
             RELEASE WINDOW ALL
          end

Return
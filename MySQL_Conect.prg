/*
 * Funções e procedimentos para conexão com MySQL
 *
*/

#include <hmg.ch>

/* Executa uma conexão ao servidor de BD MySQL local ou na web */

Function ConectaMySQL()

         MsgStatus( 'Conectando "' + g_oHostConect:DataBase + '"...', 'dbRefresh' )

         DesconectaBD()

         MemVar->g_oServer := TMySQLServer():New( g_oHostConect:Name, g_oHostConect:User, g_oHostConect:Password )

         if (MemVar->g_oServer == NIL) .or. MemVar->g_oServer:NetErr()

            DesconectaBD()
            SysWait()

            MemVar->g_oServer := TMySQLServer():New( g_oHostConect:Name, g_oHostConect:User, g_oHostConect:Password )

            if (MemVar->g_oServer == NIL) .or. MemVar->g_oServer:NetErr()
               if !(MemVar->g_oServer == NIL)
                  RegistraLog( 'Sem conexão com a internet. Erro de autenticação servidor MySQL (' + g_oHostConect:Name + '): ' + g_oServer:Error() )
               end
               PlayAsterisk()
               //MsgStop('Erro de conexão com servidor MySQL:' + CRLF + MemVar->g_oServer:Error(), 'Servidor: ' + g_oHostConect:Name)
               MsgStatus( 'Erro de conexão com servidor MySQL "' + g_oHostConect:Name + '"', 'dbError' )
               DesconectaBD()
               Return .F.
            end
         end

         /* Servidor de Banco de dados MySQL conectado, Conecta a uma específica Base De Dados */
         MemVar->g_oServer:SelectDB( MemVar->g_oHostConect:DataBase )

         if MemVar->g_oServer:NetErr()
            PlayAsterisk()
            //MsgStop('Erro de conexão com a base de dados ' + MemVar->g_oHostConect:DataBase + '.' + CRLF + MemVar->g_oServer:Error(), 'Servidor: ' + g_oHostConect:Name)
            MsgStatus( 'Erro de conexão com a base de dados "' + g_oHostConect:DataBase + '"' )
            DesconectaBD()
            Return .F.
         end

         g_oHostConect:Conectado := .T.

         MsgStatus()

Return .T.
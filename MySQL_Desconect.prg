/*
 * Funções e procedimentos para conexão com MySQL
 *
*/

/* Executa uma desconexão ao servidor de BD MySQL local ou na web */

Procedure DesconectaMySQL
          Local cSQL, oQuery

          if ( g_oHostConect:Conectado )
             // Libera usuário logado no sistema desta conexão aberta que será fechada
             cSQL   := "UPDATE usuarios SET user_conect_id = 0 WHERE user_id = 3;"
             oQuery := ExecutaQuery(cSQL)
             if !(oQuery == NIL)
                oQuery:Destroy()
                oQuery := NIL
             end
          end

          DesconectaBD()

Return

Procedure DesconectaBD()

          // Este procedimento também é usado em outras funções, não é exclusiva de DesconectaMySQL()

         g_oHostConect:Conectado := .F.

          if !(MemVar->g_oServer == NIL)
             MemVar->g_oServer:Destroy()
             MemVar->g_oServer := NIL
          end

Return
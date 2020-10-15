/*
 * Executa uma query junto ao banco de dados para conexão com MySQL
 *
*/

#include <hmg.ch>

declare window Main

Function ExecutaQuery( cSQL )

         Local cSQLx
         Local oQuery

         if !(g_oHostConect:Conectado)
            if !ConectaMySQL()
               Return (oQuery)
            end
         end

         /*
          * Para efeito de TRIGGERS
          * Se não for SELECT, pode ser INSERT, UPDATE ou DELETE, então seta o id da conexão para usuário logado
          */

         if !( HMG_UPPER(HB_ULEFT(cSQL,7)) == "SELECT " ) .and. !("UPDATE usuarios SET user_conect_id = 0" $ cSQL)

            // Reseta usuários conectados a mais de 8 horas, isso ocorre qdo este sistema é abortado pelo usuário ou por erro de execução deixando conexão em aberto
            cSQLx  := "UPDATE usuarios SET user_conect_id = 0, user_conected_em = NULL WHERE TIMEDIFF(NOW(), user_conected_em) > TIME('08:00:00');"
            oQuery := g_oServer:Query(cSQLx)

            if !(oQuery == NIL)
               oQuery:Destroy()
            end

            oQuery := NIL

            // Atrubui ID da conexão atual ao usuário ativo
            cSQLx  := "UPDATE usuarios SET user_conect_id = CONNECTION_ID(), user_conected_em = NOW() WHERE user_id = 3;"
            oQuery := g_oServer:Query(cSQLx)

            if (oQuery == NIL) .or. oQuery:NetErr()
               if !ConectaMySQL()
                  Return (oQuery)
               end
               oQuery := g_oServer:Query(cSQLx)
            end

            if Servidor_Ocupado(oQuery)
               oQuery:Destroy()
               SysWait(1)
               oQuery := g_oServer:Query(cSQLx)
            end

            oQuery:Destroy()
            SysWait(1)

         end

         oQuery := g_oServer:Query( cSQL )

         if (oQuery == NIL) .or. oQuery:NetErr()
            if !ConectaMySQL()
               Return (oQuery)
            end
            oQuery := g_oServer:Query( cSQL )
         end

         if Servidor_Ocupado(oQuery)
            oQuery:Destroy()
            SysWait(1)
            oQuery := g_oServer:Query( cSQL )
         end

Return ( oQuery )

Function Servidor_Ocupado(oQry)
Return (oQry:NetErr() .and. 'server has gone away' $ oQry:Error())

Function ExecutouQuery(oBjeto, _sql)
         Local lExecutou := (ValType(oBjeto) == "O")
         Local n

         if !lExecutou

            MsgStatus( 'Conexão interrompida, tentando reconectar...', 'dbOff' )

            FOR n := 1 to 10

               SysWait(1)

               // Variável "oBjeto" passado por referência, em caso de sucesso na conexão a rotina mãe terá o objeto restaurado.

               oBjeto := ExecutaQuery(_sql)

               if ( lExecutou := (ValType(oBjeto) == "O") )
                  g_oHostConect:Conectado := .T.
                  MsgStatus()
                  EXIT
               end

            NEXT n

         End

Return (lExecutou)

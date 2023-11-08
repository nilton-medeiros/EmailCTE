#include <hmg.ch>
#include "hbclass.ch"
#define true  .T.
#define false .F.


CREATE CLASS HostConect

     VAR Name      INIT ''
     VAR User      INIT ''
     VAR DataBase  INIT ''
     VAR Password  INIT ''
     VAR Conectado INIT false

END CLASS

CREATE CLASS InfoEmpresa

     VAR Empresa         INIT {}
     VAR nPos            INIT 0
     VAR id              INIT 0
     VAR Ambiente        INIT 0
     VAR Nome            INIT ''
     VAR Fantasia        INIT ''
     VAR sigla_cia       INIT ''
     VAR CNPJ            INIT ''
     VAR Fone            INIT ''

     VAR email_contabil  INIT ''
     VAR email_comercial INIT ''
     VAR Portal          INIT ''
     VAR gmail1_login    INIT ''
     VAR gmail1_senha    INIT ''
     VAR gmail2_login    INIT ''
     VAR gmail2_senha    INIT ''

     VAR smtp_servidor   INIT ''
     VAR smtp_pass       INIT ''
     VAR smtp_email      INIT ''
     VAR smtp_login      INIT ''
     VAR smtp_senha      INIT ''
     VAR smtp_porta      INIT 0
     VAR smtp_autentica  INIT 0
     VAR email_CCO       INIT ''

     VAR pathShared INIT ''
     VAR cte_path  INIT ''

     METHOD Adds()
     METHOD Clean()
     METHOD QtdeEmpresas()
     METHOD SetFolders(n) SETGET
     METHOD SetEmpresa() SETGET
     METHOD SetByID() SETGET

END CLASS

METHOD Adds( oBjetoEmp ) CLASS InfoEmpresa
     AADD( ::Empresa, oBjetoEmp )
     ::nPos := HMG_LEN(::Empresa)
RETURN SELF

METHOD Clean() CLASS InfoEmpresa

     ::Empresa         := {}
     ::nPos            := 0
     ::id              := 0
     ::Ambiente        := 0
     ::Nome            := ''
     ::Fantasia        := ''
     ::sigla_cia       := ''
     ::CNPJ            := ''
     ::Fone            := ''

     ::email_contabil  := ''
     ::email_comercial := ''
     ::Portal          := ''
     ::gmail1_login    := ''
     ::gmail1_senha    := ''
     ::gmail2_login    := ''
     ::gmail2_senha    := ''

     ::smtp_servidor   := ''
     ::smtp_pass       := ''
     ::smtp_email      := ''
     ::smtp_login      := ''
     ::smtp_senha      := ''
     ::smtp_porta      := 0
     ::smtp_autentica  := 0
     ::email_CCO       := ''

     ::pathShared := ''
     ::cte_path  := ''

RETURN SELF

METHOD QtdeEmpresas() CLASS InfoEmpresa
RETURN HMG_LEN( ::Empresa )

METHOD SetFolders(n) CLASS InfoEmpresa
       Local nLen := HMG_LEN(::Empresa)

       if !( ValType(n) == 'N' ) .OR. ( n < 1 ) .OR. ( n > nLen )
          n := 1
       end

       if ( nLen == 0 )

          ::nPos            := 0
          ::id              := 0
          ::Ambiente        := 0
          ::Nome            := ''
          ::Fantasia        := ''
          ::sigla_cia       := ''
          ::CNPJ            := ''
          ::Fone            := ''

          ::email_contabil  := ''
          ::email_comercial := ''
          ::Portal          := ''
          ::gmail1_login    := ''
          ::gmail1_senha    := ''
          ::gmail2_login    := ''
          ::gmail2_senha    := ''

          ::smtp_servidor   := ''
          ::smtp_pass       := ''
          ::smtp_email      := ''
          ::smtp_login      := ''
          ::smtp_senha      := ''
          ::smtp_porta      := 0
          ::smtp_autentica  := 0
          ::email_CCO       := ''

          ::pathShared  := ''
          ::cte_path  := ''

       else

          ::nPos            := n
          ::id              := ::Empresa[n]:FieldGet('emp_id')
          ::Ambiente        := ::Empresa[n]:FieldGet('emp_ambiente_sefaz')
          ::Nome            := AllTrim( ::Empresa[n]:FieldGet('empresa') )
          ::Fantasia        := AllTrim( ::Empresa[n]:FieldGet('emp_nome_fantasia') )
          ::sigla_cia       := AllTrim( ::Empresa[n]:FieldGet('emp_sigla_cia') )
          ::CNPJ            := ::Empresa[n]:FieldGet('emp_cnpj')
          ::Fone            := AllTrim( ::Empresa[n]:FieldGet('emp_fone1') )
          ::email_contabil  := AllTrim( ::Empresa[n]:FieldGet('emp_email_contabil') )
          ::email_comercial := AllTrim( ::Empresa[n]:FieldGet('emp_email_comercial') )
          ::Portal          := AllTrim( ::Empresa[n]:FieldGet('emp_portal') )
          ::gmail1_login    := AllTrim( ::Empresa[n]:FieldGet('emp_gmail1_login') )
          ::gmail1_senha    := AllTrim( ::Empresa[n]:FieldGet('emp_gmail1_senha') )
          ::gmail2_login    := AllTrim( ::Empresa[n]:FieldGet('emp_gmail2_login') )
          ::gmail2_senha    := AllTrim( ::Empresa[n]:FieldGet('emp_gmail2_senha') )

          ::smtp_servidor   := AllTrim( ::Empresa[n]:FieldGet('emp_smtp_servidor') )
          ::smtp_pass       := AllTrim( ::Empresa[n]:FieldGet('emp_smtp_pass') )
          ::smtp_email      := AllTrim( ::Empresa[n]:FieldGet('emp_smtp_email') )
          ::smtp_login      := AllTrim( ::Empresa[n]:FieldGet('emp_smtp_login') )
          ::smtp_senha      := AllTrim( ::Empresa[n]:FieldGet('emp_smtp_senha') )

          if Empty(::smtp_pass)
            ::smtp_pass := ::smtp_senha
          endif

          ::smtp_porta      := ::Empresa[n]:FieldGet('emp_smtp_porta')
          ::smtp_autentica  := ::Empresa[n]:FieldGet('emp_smtp_autentica')
          ::email_CCO       := ::Empresa[n]:FieldGet('emp_email_CCO')

          ::pathShared := RegistryRead('HKEY_CURRENT_USER\SOFTWARE\Sistrom\DFeMonitor\InstallPath\dfePath')
          ::cte_path := ::pathShared + ::Empresa[::nPos]:FieldGet('emp_cnpj') + "\CTe\"

          if !hb_DirExists( ::cte_path )
            hb_DirBuild( ::cte_path )
          end

       end

RETURN SELF

METHOD SetEmpresa( n ) CLASS InfoEmpresa
       Local nLen := HMG_LEN(::Empresa)

       if !( ValType(n) == 'N' ) .OR. ( n < 1 ) .OR. ( n > nLen )
          n := 1
       end

       if ( nLen == 0 )

          ::nPos            := 0
          ::id              := 0
          ::Ambiente        := 0
          ::Nome            := ''
          ::Fantasia        := ''
          ::sigla_cia       := ''
          ::CNPJ            := ''
          ::Fone            := ''

          ::email_contabil  := ''
          ::email_comercial := ''
          ::Portal          := ''
          ::gmail1_login    := ''
          ::gmail1_senha    := ''
          ::gmail2_login    := ''
          ::gmail2_senha    := ''

          ::smtp_servidor   := ''
          ::smtp_pass       := ''
          ::smtp_email      := ''
          ::smtp_login      := ''
          ::smtp_senha      := ''
          ::smtp_porta      := 0
          ::smtp_autentica  := 0
          ::email_CCO       := ''

          ::cte_path  := ''

       else

          ::nPos            := n
          ::id              := ::Empresa[n]:FieldGet('emp_id')
          ::Ambiente        := ::Empresa[n]:FieldGet('emp_ambiente_sefaz')
          ::Nome            := AllTrim( ::Empresa[n]:FieldGet('empresa') )
          ::Fantasia        := AllTrim( ::Empresa[n]:FieldGet('emp_nome_fantasia') )
          ::sigla_cia       := AllTrim( ::Empresa[n]:FieldGet('emp_sigla_cia') )
          ::CNPJ            := ::Empresa[n]:FieldGet('emp_cnpj')
          ::Fone            := AllTrim( ::Empresa[n]:FieldGet('emp_fone1') )
          ::email_contabil  := AllTrim( ::Empresa[n]:FieldGet('emp_email_contabil') )
          ::email_comercial := AllTrim( ::Empresa[n]:FieldGet('emp_email_comercial') )
          ::Portal          := AllTrim( ::Empresa[n]:FieldGet('emp_portal') )
          ::gmail1_login    := AllTrim( ::Empresa[n]:FieldGet('emp_gmail1_login') )
          ::gmail1_senha    := AllTrim( ::Empresa[n]:FieldGet('emp_gmail1_senha') )
          ::gmail2_login    := AllTrim( ::Empresa[n]:FieldGet('emp_gmail2_login') )
          ::gmail2_senha    := AllTrim( ::Empresa[n]:FieldGet('emp_gmail2_senha') )

          ::smtp_servidor   := AllTrim( ::Empresa[n]:FieldGet('emp_smtp_servidor') )
          ::smtp_pass       := AllTrim( ::Empresa[n]:FieldGet('emp_smtp_pass') )
          ::smtp_email      := AllTrim( ::Empresa[n]:FieldGet('emp_smtp_email') )
          ::smtp_login      := AllTrim( ::Empresa[n]:FieldGet('emp_smtp_login') )
          ::smtp_senha      := AllTrim( ::Empresa[n]:FieldGet('emp_smtp_senha') )
          ::smtp_porta      := ::Empresa[n]:FieldGet('emp_smtp_porta')
          ::smtp_autentica  := ::Empresa[n]:FieldGet('emp_smtp_autentica')
          ::email_CCO       := ::Empresa[n]:FieldGet('emp_email_CCO')

          ::cte_path := ::pathShared + ::Empresa[n]:FieldGet('emp_cnpj') + "\CTe\"

       end

RETURN SELF


METHOD SetByID( nId ) CLASS InfoEmpresa

       ::nPos := hb_Ascan( ::Empresa, {|oVal| oVal:FieldGet('emp_id') == nId} )

       if ( ::nPos == 0 )

          ::id              := 0
          ::Ambiente        := 0
          ::Nome            := ''
          ::Fantasia        := ''
          ::sigla_cia       := ''
          ::CNPJ            := ''
          ::Fone            := ''

          ::email_contabil  := ''
          ::email_comercial := ''
          ::Portal          := ''
          ::gmail1_login    := ''
          ::gmail1_senha    := ''
          ::gmail2_login    := ''
          ::gmail2_senha    := ''

          ::smtp_servidor   := ''
          ::smtp_pass       := ''
          ::smtp_email      := ''
          ::smtp_login      := ''
          ::smtp_senha      := ''
          ::smtp_porta      := 0
          ::smtp_autentica  := 0
          ::email_CCO       := ''

          ::pathShared := ''
          ::cte_path  := ''

       else

          ::id              := ::Empresa[::nPos]:FieldGet('emp_id')
          ::Ambiente        := ::Empresa[::nPos]:FieldGet('emp_ambiente_sefaz')
          ::Nome            := AllTrim( ::Empresa[::nPos]:FieldGet('empresa') )
          ::Fantasia        := AllTrim( ::Empresa[::nPos]:FieldGet('emp_nome_fantasia') )
          ::sigla_cia       := AllTrim( ::Empresa[::nPos]:FieldGet('emp_sigla_cia') )
          ::CNPJ            := ::Empresa[::nPos]:FieldGet('emp_cnpj')
          ::Fone            := AllTrim( ::Empresa[::nPos]:FieldGet('emp_fone1') )
          ::email_contabil  := AllTrim( ::Empresa[::nPos]:FieldGet('emp_email_contabil') )
          ::email_comercial := AllTrim( ::Empresa[::nPos]:FieldGet('emp_email_comercial') )
          ::Portal          := AllTrim( ::Empresa[::nPos]:FieldGet('emp_portal') )
          ::gmail1_login    := AllTrim( ::Empresa[::nPos]:FieldGet('emp_gmail1_login') )
          ::gmail1_senha    := AllTrim( ::Empresa[::nPos]:FieldGet('emp_gmail1_senha') )
          ::gmail2_login    := AllTrim( ::Empresa[::nPos]:FieldGet('emp_gmail2_login') )
          ::gmail2_senha    := AllTrim( ::Empresa[::nPos]:FieldGet('emp_gmail2_senha') )

          ::smtp_servidor   := AllTrim( ::Empresa[::nPos]:FieldGet('emp_smtp_servidor') )
          ::smtp_pass       := AllTrim( ::Empresa[::nPos]:FieldGet('emp_smtp_pass') )
          ::smtp_email      := AllTrim( ::Empresa[::nPos]:FieldGet('emp_smtp_email') )
          ::smtp_login      := AllTrim( ::Empresa[::nPos]:FieldGet('emp_smtp_login') )
          ::smtp_senha      := AllTrim( ::Empresa[::nPos]:FieldGet('emp_smtp_senha') )
          ::smtp_porta      := ::Empresa[::nPos]:FieldGet('emp_smtp_porta')
          ::smtp_autentica  := ::Empresa[::nPos]:FieldGet('emp_smtp_autentica')
          ::email_CCO       := ::Empresa[::nPos]:FieldGet('emp_email_CCO')

          ::cte_path := ::pathShared + ::Empresa[::nPos]:FieldGet('emp_cnpj') + "\CTe\"

       end

RETURN SELF

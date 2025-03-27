Class Ajustes{
[int[]]$posto
[String]$nomeFormatado
[String]$nickname
[String]$firstName
[String]$surname

[String]$upn
[String]$CPF
[String]$employeeID
[String]$datAdm
[String]$datNas

[String[]]$memberOf
[String]$pathOU
[String]$office
[String]$pager

[String]$jobTitle
[String]$department
[String]$empresa
[String]$gestorOu
[String]$emailGestor
[String]$employeeIDGestor
[String]$gestorName
[String]$domain
[String]$defaultOu
[String]$password
[bool]$ramal


Ajustes($UsersAdmitidos){
        $this.employeeID = $UsersAdmitidos.employeeID
        $this.employeeIDGestor = $UsersAdmitidos.employeeIDGestor
        $this.datNas = $UsersAdmitidos.nasc
        $this.datAdm = $UsersAdmitidos.admiss
        $this.CPF = $UsersAdmitidos.CPF
        $this.jobtitle = $UsersAdmitidos.cargo
        $this.posto= $this.TratarCODPosto($UsersAdmitidos.posto)
        $this.TratarUser($this.posto, $UsersAdmitidos.name, $UsersAdmitidos)

    }

    
[void]TratarUser($posto, $name, $User){
    $this.AjusteLogin($name)
    $UserExistente = Get-ADUser -Filter {extensionAttribute3 -eq $this.CPF}
    if(!$this.nickname -and !$UserExistente){
    EnviarEmail $User "manual" "eduardo.mendes@dufrio.com.br"
    }
    
    $this.password = $this.NewPassword(9,5,1,1)

    $emp = $this.posto[0]
    $filial = $this.posto[1]
    $departamento = $this.posto[2]
    $dados = $this.validarFilial($emp,$filial,$departamento, $this.jobTitle)
    
      #Valida o path, se não existir, envia para ou padrão
      $ou = "OU=$($dados.SiglaOUFilial)_$($dados.ouPathSetor),OU=Users_$($dados.SiglaOUFilial),OU=Users Dufrio,DC=dufrio,DC=local"
      if($this.ValidateOUPath($ou)){ $ou = "OU=$($this.defaultOu)_New,OU=Users_$($this.defaultOu),OU=Users Dufrio,DC=dufrio,DC=local"}


      #Atribui os membersOF a partir de uma função
      if(!$UserExistente){$this.upn = "$($this.nickname)$($this.domain)"}
      $this.pathOU = $ou
      
      $this.GetGestor($this.employeeIDGestor)

      $this.office = $dados.office
      $this.department = $dados.departamento
      $this.pager = $dados.pagerSigla
      $this.GetMembersOf($dados.siglaMembers,$dados.departamento)
    
  }
[int[]]TratarCODPosto([String]$CodPosto){
        # Faz o split da string usando o ponto como separador
        $parts = $CodPosto -split '\.'

        # Inicializa um array vazio para armazenar os inteiros
        $integers = @()

        # Para cada parte da string, converte em inteiro e adiciona ao array
        foreach ($part in $parts) {
            $integers += [int]$part
        }
        return $integers
        
    }
[String]NewPassword([int]$length,[int]$upper=1,[int]$lower=1, [int]$numeric){
        # Validação manual do parâmetro $length

        if ($length -lt 4) {
            throw "O comprimento da senha deve ser maior ou igual a 4."
        }

        # Valida se a soma de upper, lower e numeric não excede o comprimento da senha
        if ($upper + $lower + $numeric -gt $length) {
            throw "O número de caracteres maiúsculos, minúsculos e numéricos deve ser menor ou igual ao comprimento da senha."
        }

        # Define os conjuntos de caracteres
        $uCharSet = "ABCDEFGHIJKLMNOPQRSTUVXJ"
        $lCharSet = "abcdefghjkmnopqrstuvxz"
        $nCharSet = "0123456789"
        
        # Inicializa o conjunto de caracteres permitido com base nos parâmetros
        $charSet = ""
        if ($upper -gt 0) { $charSet += $uCharSet }
        if ($lower -gt 0) { $charSet += $lCharSet }
        if ($numeric -gt 0) { $charSet += $nCharSet }

        # Converte o conjunto de caracteres para um array
        $charSet = $charSet.ToCharArray()

        # Gera os bytes de forma aleatória
        $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
        $bytes = New-Object byte[]($length)
        $rng.GetBytes($bytes)

        # Gera a senha inicial a partir do conjunto de caracteres aleatórios
        $result = New-Object char[]($length)
        for ($i = 0 ; $i -lt $length ; $i++) {
            $result[$i] = $charSet[$bytes[$i] % $charSet.Length]
        }

        $PASSWD = (-join $result)
        $passwordKey = "DU@" + $PASSWD

        # Verifica se o número de caracteres maiúsculos, minúsculos e numéricos está correto
        $valid = $true
        if ($upper   -gt ($passwordKey.ToCharArray() | Where-Object { $_ -cin $uCharSet.ToCharArray() }).Count) { $valid = $false }
        if ($lower   -gt ($passwordKey.ToCharArray() | Where-Object { $_ -cin $lCharSet.ToCharArray() }).Count) { $valid = $false }
        if ($numeric -gt ($passwordKey.ToCharArray() | Where-Object { $_ -cin $nCharSet.ToCharArray() }).Count) { $valid = $false }

        # Caso a senha gerada não seja válida, chama o método novamente
        if (!$valid) {
            return $this.NewPassword($length, $upper, $lower, $numeric)
        }
        else {return $passwordKey}

    }
[PSCustomObject]ValidarFilial([int]$emp, [int]$filial, [int]$departamento, [String]$cargo){
    [String]$descricao,$sigla,$siglaOUFilial,$pagerSigla,$siglaMembers =""
    switch($emp){
        
        1{
        # EMPRESA 1
            switch($filial){
                1{
                    #TABELA DE FILIAIS
                }
            }            
} 
        5{
            #EMPRESA 5
        }
        default{
        $descricao = "CODIGO DE EMPRESA NAO CAD:EMP"
        $sigla = "NULL"
        }
    }
    $siglaOUFilial = $sigla
    if(!$siglaMembers){$siglaMembers = $sigla} 
    if(!$pagerSigla){$pagerSigla = $sigla}
    $departamentoObject = $this.ValidarDepartamento([int]$emp,[int]$filial,[int]$departamento, [String]$cargo)
    return [PSCustomObject]@{
        office        = $descricao
        SiglaOUFilial = $siglaOUFilial
        pagerSigla = $pagerSigla
        siglaMembers = $siglaMembers
        departamento = $departamentoObject.descricao
        ouPathSetor = $departamentoObject.OuPath
    }
}
[PSCustomObject]ValidarDepartamento([int]$emp,[int]$filial, [int]$departament, [String]$cargo){
$descricao = ""
$ouPath =  ""
switch($emp){
 
    1{
    switch($departament){
            1   { $descricao = "ABASTECIMENTO CD"}
            2   { $descricao = "ABASTECIMENTO LOJA"; $ouPath = "COMPRAS" }
            3   { $descricao = "ADMINISTRACAO DE PESSOAL"; $ouPath = "RH" }
            4   { $descricao = "ADMINISTRATIVO CANAIS DIGITAIS"; $ouPath = "eCommerce" }
            5   { $descricao = "ADMINISTRATIVO FILIAL";if($cargo -like "*CAIXA*"){$ouPath = "Caixa"}else{$ouPath = "Administrativo"}}
            6   { $descricao = "ATENDIMENTO"; $ouPath = "Caixa" }
            7   { $descricao = "COBRANCA"; $ouPath = "Cobranca" }
            8   { $descricao = "COMERCIAL"}
            9   { $descricao = "CONSELHO"}
            10  { $descricao = "CONTAS A PAGAR"; $ouPath = "Contas_A_Pagar" }
            11  { $descricao = "CONTAS A RECEBER"; $ouPath = "Cobranca" }
            12  { $descricao = "CONTROLADORIA"; $ouPath = "Controladoria" }
            13  { $descricao = "CREDITO E CADASTRO"; $ouPath = "Credito" }
            14  { $descricao = "CSD - SERVICOS ADMINISTRATIVOS"; $ouPath = "Administrativo" }
            15  { $descricao = "DESENVOLVIMENTO DE PESSOAS"; $ouPath = "RH" }
            16  { $descricao = "DESENVOLVIMENTO DE PRODUTOS"}
            17  { $descricao = "DIRETORIA COMERCIAL"; $ouPath = "Diretoria" }
            18  { $descricao = "DIRETORIA DE CONTROLADORIA"; $ouPath = "Diretoria" }
            19  { $descricao = "DIRETORIA FINANCEIRA"; $ouPath = "Diretoria" }
            20  { $descricao = "DIRETORIA DE SUPPLY CHAIN"; $ouPath = "Diretoria" }
            21  { $descricao = "DIRETORIA DE TI E RH"; $ouPath = "Diretoria" }
            22  { $descricao = "FINANCEIRO"; $ouPath = "Tesouraria" }
            23  { $descricao = "FISCAL"; $ouPath = "Contabilidade" }
            24  { $descricao = "IMPORTACAO"; $ouPath = "COMPRAS" }
            25  { $descricao = "JURIDICO"; $ouPath = "Juridico" }
            26  { $descricao = "LOGISTICA CD"; $ouPath = "Logistica" }
            27  { $descricao = "LOGISTICA LOJA"; $ouPath = "Logistica" }
            28  { $descricao = "REFORMA, OBRAS E EXPANSAO"; $ouPath = "Administrativo" }
            29  { $descricao = "MARKETING"; $ouPath = "Marketing" }
            30  { $descricao = "MARKETPLACE"; $ouPath = "eCommerce" }
            31  { $descricao = "PLANEJAMENTO"; $ouPath = "Gerentes" }
            32  { $descricao = "PLANEJAMENTO FINANCEIRO"; $ouPath = "Gerentes" }
            33  { $descricao = "PLATAFORMA DE NEGOCIOS"}
            34  { $descricao = "POS VENDAS"; $ouPath = "Garantia" }
            35  { $descricao = "PRE VENDA"}
            36  { $descricao = "PRESIDENCIA"}
            37  { $descricao = "PROCESSOS E PROJETOS"}
            38  { $descricao = "SAC ATENDIMENTO"; $ouPath = "SAC" }
            39  { $descricao = "SAC BACKOFFICE"; $ouPath = "SAC" }
            40  { $descricao = "SITE - ECOMMERCE"; $ouPath = "eCommerce" }
            41  { $descricao = "SSMA - SEGURANCA, SAUDE, MEIO AMBIENTE";
                if($filial -eq 01){$ouPath= "RH"} else{ $ouPath="Administracao"}           
            }
            42  { $descricao = "SUPPLY CHAIN"}
            43  { $descricao = "TELEVENDAS"; $ouPath = "Vendas" }
            44  { $descricao = "TI OPERACOES"; $ouPath = "TI" }
            45  { $descricao = "TI PLANEJAMENTO E CONTROLE"; $ouPath = "TI" }
            46  { $descricao = "TRANSPORTES"; $ouPath = "Logistica" }
            47  { $descricao = "VENDAS"; $ouPath = "Vendas" }
            48  { $descricao = "VENDAS CAMARAS FRIAS"; $ouPath = "Vendas" }
            49  { $descricao = "VENDAS CLIMATIZACAO"; $ouPath = "Vendas" }
            50  { $descricao = "VENDAS CORPORACOES"; $ouPath = "Vendas" }
            51  { $descricao = "VENDAS ENERGIA RENOVAVEL"; $ouPath = "Vendas" }
            52  { $descricao = "VENDAS EXPRESS"; $ouPath = "Vendas" }
            53  { $descricao = "VENDAS REVENDAS"; $ouPath = "Vendas" }
            54  { $descricao = "ENDOMARKETING"; $ouPath = "RH" }
            55  { $descricao = "FAMILY OFFICE"; $ouPath = "Tesouraria" }
            56  { $descricao = "PAO DOS POBRES"; $ouPath = "RH" }
            57  { $descricao = "MESA DE NEGOCIACAO"}
            59  { $descricao = "CADASTRO"; $ouPath = "Compras" }
            60  { $descricao = "INTELIGENCIA DE NEGOCIO"; $ouPath = "Controladoria" }
            61  { $descricao = "PLANEJAMENTO COMERCIAL"}
            62  { $descricao = "COMPRAS"; $ouPath = "Compras" }
            63  { $descricao = "INVENTARIO"; $ouPath = "Logistica" }
            64  { $descricao = "ADMINISTRATIVO"; $ouPath = "Administrativo" }
            65  { $descricao = "ATRACAO E SELECAO"; $ouPath = "RH" }
            66  { $descricao = "CONSULTORIA DE RH"; $ouPath = "RH" }
            67  { $descricao = "TECNICO COMERCIAL"; $ouPath = "Vendas" }
            68  { $descricao = "DIRETORIA DE SOLUCOES E SERVICOS"}
            69  { $descricao = "ENGENHARIA DE SOLUCOES E SERVICOS"}
            70  { $descricao = "SERVICOS"}
            71  { $descricao = "OPERACOES COMERCIAIS"}
            72  { $descricao = "COMPLIANCE" }
            73  { $descricao = "CONTABILIDADE"; $ouPath = "Contabilidade" }
            74  { $descricao = "AUDITORIA INTERNA"; $ouPath = "Juridico" }
            75  { $descricao = "CUSTOS E ESTOQUE"; $ouPath = "Administrativo" }
            76  { $descricao = "TI ENGENHARIA E CIENCIA DE DADOS"; $ouPath = "TI" }
            77  { $descricao = "PROJETOS"; $ouPath = "Administrativo" }
            100 { $descricao = "AFASTADOS INSS"}
            247 { $descricao = "Vendas Co-working"}
            default { $descricao = "CODIGO NAO CADASTRADO:SCRIPT";$ouPath="New" }
            }
        }
    5{
            switch($departament){
                1  { $descricao = "ADMINISTRATIVO";$ouPath = "ADMINISTRATIVO" }
                2  { $descricao = "CONFORMACAO"; $ouPath = "PRODUCAO" }
                3  { $descricao = "ESTAMPARIA"; $ouPath = "PRODUCAO" }
                4  { $descricao = "INDUSTRIAL"; $ouPath = "ADMINISTRATIVO" }
                5  { $descricao = "LOGISTICA"; $ouPath = "LOGISTICA" }
                6  { $descricao = "MANUTENCAO"; $ouPath = "" }
                7  { $descricao = "MONTAGEM"; $ouPath = "PRODUCAO" }
                8  { $descricao = "QUALIDADE"; $ouPath = "PRODUCAO" }
                9  { $descricao = "USINAGEM"; $ouPath = "PRODUCAO" }
                10 { $descricao = "ELETRONICO"; $ouPath = "PRODUCAO" }
                11 { $descricao = "SSMA - SEGURANCA, SAUDE, MEIO AMBIENTE"; $ouPath = "ADMINISTRATIVO" }
                12 { $descricao = "PRODUCAO"; $ouPath = "PRODUCAO" }
                13 { $descricao = "PROJETOS"; $ouPath = "PROJETOS" }
                14 { $descricao = "DIRETORIA COMERCIAL"; $ouPath = "" }
                15 { $descricao = "COMPRAS"; $ouPath = "LOGISTICA" }
                16 { $descricao = "CAMARA FRIAS"; $ouPath = "LOGISTICA" }
                17 { $descricao = "PRESIDENCIA"; $ouPath = "" }
                18 { $descricao = "VENDAS"; $ouPath = "VENDAS" }
                default { $descricao = "CÓDIGO NÃO CADASTRADO"; $ouPath = "New" }

            }


        }
    default{
    $descricao = "Código de Empresa não existente:DEP"
    $SIGLA = "NULL"

    }
    }
return [PSCustomObject]@{
        Descricao                = $descricao
        OuPath = $ouPath
        #validar por if
    }
}
[Boolean]ValidateOUPath([String]$ouPath) {
    
    try {
        # Tenta obter a OU com base no caminho fornecido
        $ou = Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ouPath'" -ErrorAction Stop
        
        if ($ou) {
            return $false
        } else {
            return $true
        }
    }
    catch {
        # Se houver um erro (ou a OU não existir), retorna falso
        return $false
    }
}
[void] GetMembersOf ([String]$filialMemberOf,[String]$departament){
    $this.memberOf = @(
        "Lista - Todos",
        "Senior",
        "SEC - Radius Users",
        "OTRS-Clientes",
        "Lista - Todos $filialMemberOf"
    )
    if($departament -like "*VENDAS*"){
        $this.memberOf += ("ACL_304-Vendas","IMP_$filialMemberOf-Vendas", "Lista - Vendedores $filialMemberOf")
        $this.ramal=$true 
    }
    elseif($departament -like "*LOGISTICA*"){
        $this.memberOf += ("ACL_305-Logistica","IMP_$filialMemberOf-Logistica", "Lista - Logistica $filialMemberOf")
    }
    elseif($departament -like "*SAC*"){
        $this.memberOf += ("ACL_322-Telefonistas","IMP_$filialMemberOf-SAC", "Lista - SAC $filialMemberOf", "Lista - SAC", "Intelipost")
    }
    elseif($departament -like "*TI OPERACOES*"){
        $this.memberOf += ("ACL_303-TI","IMP_$filialMemberOf-TI")
    }
    if($departament -like "*MARKETING*"){
        $this.memberOf += ("ACL_308-Marketing","IMP_POA-MARKETING-COLOR")
    }
    if($departament -like "*RH*"){
        $this.memberOf += ("ACL_306-RH","IMP_$filialMemberOf-RH", "Lista - RH $filialMemberOf")
    }
    if($departament -like "*ADMINISTRATIVO*"){
        $this.memberOf += ("ACL_307-Administrativo","IMP_$filialMemberOf-Administrativo")
    }
}
[void]AjusteLogin($nameRaw){

    $this.nomeFormatado = $this.FormatName($nameRaw)
    $this.nickname = $this.CreateNickname($this.nomeFormatado)
    $fullnameSplit = $this.nomeFormatado -split " ",2
    $this.firstName = $fullnameSplit[0]
    $this.surname = $fullnameSplit[1]
    
}
[String]FormatName($completeName){
        $completeName = $completeName -Replace("ç","c")`
        -Replace("ã","a")`
        -Replace("á","a")`
        -Replace("â","a")`
        -Replace("õ","o")`
        -Replace("ó","o")`
        -Replace("ô","o")`
        -Replace("ú","u")`
        -Replace("é","e")`
        -Replace("ê","e")`
        -Replace("Ç","c")`
        -Replace("Á","a")`
        -Replace("Ã","a")`
        -Replace("Â","a")`
        -Replace("Õ","o")`
        -Replace("Ó","o")`
        -Replace("Ô","o")`
        -Replace("É","e")`
        -Replace("Ê","e")`
        -Replace("Ú","u")`
        -Replace("í","i")`
        -Replace("  "," ")

        $DefaultText = (Get-Culture).TextInfo
        $Name = $DefaultText.ToTitleCase($completeName.ToLower())

        return $Name
    }
[String] CreateNickname($correctName){  
        $correctName = $correctName.toLower()
        $split = $correctName -split " "
        $nameParts=@()
        $nicknames= @()
        $seen = @{}
        foreach ($campo in $split) {
            if (
                $campo -ne "da" -AND 
                $campo -ne "de" -AND 
                $campo -ne "do" -AND 
                $campo -ne "das" -AND 
                $campo -ne "dos" -AND 
                $campo -ne "e" -AND 
                $campo -ne "la" -AND 
                $campo -ne "los" -AND 
                $campo -ne "le"
                ){$nameParts += $campo}
        }
        
        foreach ($part1 in $nameParts) {
            for ($i = $nameParts.Length - 1; $i -ge 0; $i--) {
                $part2 = $nameParts[$i]
                if ($part1 -ne $part2) {
                    $combination = "$($part1).$($part2)"
                    if (-not $seen.Contains($combination)) {
                        $nicknames += $combination
                        $seen[$combination] = $true
                    }
                }
            }
        }
        $nicknameCorreto = ""
        foreach($nickname in $nicknames){

            $userExists = Get-ADUser -Filter {SamAccountName -eq $nickname}
            if(!$userExists){
                $nickname.lenght
                if($nickname.Length -le 20){
                    Write-Host $nickname
                    $nicknameCorreto = $nickname
                    return $nicknameCorreto
                }

            }
        }
        return $nicknameCorreto
        }
[void]GetGestor($employeeIDGestor){
    if(!$employeeIDGestor){
    Write-Host "Sem gestor"
    }else{
        $gestorUser = Get-ADUser -Filter {employeeid -eq $employeeIDGestor}
        $this.emailGestor = $gestorUser.UserPrincipalName
        $this.gestorOu = $gestorUser.DistinguishedName
        $this.gestorName = $gestorUser.name
    }
    }
}
Class CriarUsuario{
 [bool]$create
    CriarUsuario($UserObject){
        $this.create = $false
        $alterar = $null
        if($UserObject.nickname){
                    #Valida se ele precisa ser Editado(PROCURA PELO CPF e DOMINIO)
        if($this.validarUserExists($UserObject.domain, $UserObject.CPF)){
            Write-Host "Usuario Vai ser modificado"
            $this.create = $this.SetUser($UserObject)
            Write-Host $UserObject
            CriarLog $UserObject "ajuste"

            EnviarEmail $UserObject "RHAJUSTE" "rh.admissoes@dufrio.com.br"
            EnviarEmail $UserObject "RHAJUSTE" "eduardo.mendes@dufrio.com.br"
            EnviarEmail $UserObject "gestorAJUSTE" "eduardo.mendes@dufrio.com.br"

        } else{
            Write-Host "Usuario Vai ser criado"
            $this.create = NewUsers $UserObject
            if($UserObject.ramal){EnviarEmail $UserObject "RAMAL" "servicedesk@dufrio.com.br"}
            EnviarEmail $UserObject "RHCRIACAO" "rh.admissoes@dufrio.com.br"
            EnviarEmail $UserObject "RHCRIACAO" "eduardo.mendes@dufrio.com.br"
            CriarLog $UserObject "criacao"
            }
            #Quando criado/Modificado o usuario, sera enviado email correspondente para rh e gestor e log
            #Parametros: ObjetoUsuario para os campos do email, String para if e mudança do corpo, e para quem
        if(!$this.create){

            EnviarEmail $UserObject "manual" "eduardo.mendes@dufrio.com.br"
            CriarLog $UserObject "manual"
            #Cria o log e avisa no titulo se é alteração ou criação do 0
        
            }else {

            }
        } 
        else {
            Write-Host "Sem nickname"
        }

    }
    [bool]validarUserExists([String]$domain, $cpf){
        $UserExistente = Get-ADUser -Filter {extensionAttribute3 -eq $cpf} -ErrorAction SilentlyContinue
        if($UserExistente.UserPrincipalName -like "*$domain*" ){
            Write-Host "Usuario com mesmo dominio e CPF"
            return $true
    } else {
            Write-Host "Usuário nao encontrado" 
            return $false
        }
    }
    [bool]SetUser($UserObject){
        $UserExistente = Get-ADUser -Filter {extensionattribute3 -eq $UserObject.CPF} -Properties SamAccountName

        Write-Host "$($UserObject.EmployeeID);$($UserObject.office);$($UserObject.jobTitle);$($UserObject.department);$($UserObject.empresa)"
        
        Set-ADUser -Identity $UserExistente.SamAccountName -EmployeeID $UserObject.EmployeeID`
                -Title $UserObject.jobTitle`
                -Department $UserObject.department`
                -Company $UserObject.empresa`
                -Office $UserObject.office

        if($UserObject.gestorOu){Set-ADUser -Identity $UserExistente.SamAccountName -Manager $UserObject.gestorOu}

        Set-ADUser -Identity $UserExistente.SamAccountName -Replace @{Pager = $UserObject.pager}
        Set-ADUser -Identity $UserExistente.SamAccountName -Replace @{extensionAttribute1 = $UserObject.datAdm}
        Move-ADObject -Identity (Get-ADUser -Identity $UserExistente.SamAccountName).DistinguishedName -TargetPath $UserObject.pathOU
        if(!$UserExistente.enabled){
            Enable-ADAccount -Identity $UserExistente.SamAccountName
            Set-ADAccountPassword -Identity $UserExistente.SamAccountName -NewPassword (ConvertTo-SecureString -AsPlainText $UserObject.password -Force)
            EnviarEmail $UserObject "RAMAL" "servicedesk@dufrio.com.br"
        }
        return $true
    }
}
function EnviarEmail($UserObject, $assunto, $emailSend){
    [String]$to = $emailSend
    [String]$smtpServer = "SMTPSRV"
    [int]$smtpPort = "PORTA"
    [String]$username = "LOGIN"
    [String]$password = "SENHA"
    #From é padrão
    [String]$from = "LOGIN"
    if($assunto -like "*RHCRIACAO*"){
        [String]$subject = "Usuario Criado no AD - $($UserObject.nickname)"
        $body = "
            Usuario: $($UserObject.upn)
            EmployeeID: $($UserObject.employeeID)
            Nome: $($UserObject.nomeFormatado)
            Gestor: $($UserObject.emailGestor)
            Cargo: $($UserObject.jobTitle)
            Filial: $($UserObject.pager)
            Essa informação é um teste de script, o e-mail oficializado é via servicedesk
        "
    }
    if($assunto -like "*RHAJUSTE*"){
        $UserOld = Get-ADUser -Filter {extensionAttribute3 -eq $UserObject.CPF}
        [String]$subject = "Usuario Alterado no AD - $($UserObject.SamAccountName)"
        $body = "
            Usuario: $($UserOld.UserPrincipalName)
            EmployeeID: $($UserObject.employeeid)
            Nome: $($UserOld.Name)
            Gestor: $(if($UserGestor){$UserGestor.UserPrincipalName}else{"Sem gestor cadastrado"})
            Cargo: $($UserObject.Title)
            Filial: $($UserObject.pager)
            Essa informação é um teste de script, o e-mail oficializado é via servicedesk
        "
    }
    elseif($assunto -like "*gestorAJUSTE*"){
        $UserOld = Get-ADUser -Filter {extensionAttribute3 -eq $UserObject.CPF} 
        if($UserObject.Manager){$UserGestor = Get-ADUser -Identity $UserObject.Manager -ErrorAction SilentlyContinue}

        [String]$subject = "Usuario Modificado - $($UserOld.SamAccountName)"
        $body = "
        Se Usuario ja ativo, desconsidere a nova senha.

            Usuario: $($UserOld.UserPrincipalName)
            Senha: $($UserObject.password)
            Nome: $($UserOld.Name)
            Gestor: $(if($UserObject.gestorName){$UserObject.gestorName}else{"Sem gestor cadastrado"})
            Cargo: $($UserObject.jobTitle)
            Filial $($UserObject.pager)
            Essa informacao e um teste de script, o login oficial sera via servicedesk
        "
    
    }
     elseif($assunto -like "*gestorCRIACAO*"){
        [String]$subject = "Usuario Criado - $($UserObject.nomeFormatado)"
        $body = "
            Usuario: $($UserObject.upn)
            Senha: $($UserObject.password)
            Nome: $($UserObject.nomeFormatado)
            Gestor: $($UserObject.emailGestor)
            Cargo: $($UserObject.jobTitle)
            Filial $($UserObject.pager)
            Essa informacao e um teste de script, o login oficial sera via servicedesk
        "
    }
    elseif($assunto -like "*RAMAL*"){
        [String]$subject = "Criar RAMAL/Conferir se não Existe - $($UserObject.nomeFormatado)"
        $body = "
            Usuario: $($UserObject.upn)
            Nome: $($UserObject.nomeFormatado)
            Gestor: $($UserObject.emailGestor)
            Cargo: $($UserObject.jobTitle)
            Filial $($UserObject.pager)
        "
    }
    elseif($assunto -like "*manual*"){
       [String]$subject = "Usuario nao criado, Validar! - $($UserObject.nomeFormatado)"
    $body = "
        Usuario: $($UserObject.nomeFormatado)
        employeeID: $($UserObject.employeeID)
    " 
    }


    # Crie as credenciais
    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($username, $securePassword)

    # Envie o e-mail
    Send-MailMessage -SmtpServer $smtpServer -Port $smtpPort -UseSsl -Credential $credential -From $from -To $to -Subject $subject -Body $body
    }
function CriarLog($UserObject,$assunto){
       if($assunto -like "*criacao*"){
       $logFilePath = "C:\Util\Mendes - Scripts\Criacao de Usuario\Logs\UsuariosCriadosCSenhaLogs.txt"
       $logContent = @"
----------------------------------------
Usuário criado com sucesso
    Nome do Colaborador: $($UserObject.nomeFormatado)
    Login: $($UserObject.nickname)`
    E-mail: $($UserObject.upn)
    Senha: $($UserObject.password)
    EmployeeId: $($UserObject.employeeID)
    Data de Admissão: $($UserObject.datAdm)
    Gestor E-mail: $($UserObject.emailGestor)
    OU: $($UserObject.pathOU)
----------------------------------------
"@
    }elseif($assunto -like "*ajuste*"){
    $UserOld = Get-ADUser -Filter {extensionAttribute3 -eq $UserObject.CPF}
       $logFilePath = "C:\Util\Mendes - Scripts\Criacao de Usuario\Logs\UsuariosCriadosCSenhaLogs.txt"
       $logContent = @"
----------------------------------------
Usuário Ajustado com sucesso
    Nome do Colaborador: $($UserOld.Name)
    Login: $($UserOld.SamAccountName)`
    E-mail: $($UserOld.UserPrincipalName)
    Senha: $($UserObject.password)
    EmployeeId: $($UserObject.employeeID)
    Data de Admissão: $($UserObject.datAdm)
    Gestor E-mail: $($UserObject.emailGestor)
    OU: $($UserObject.pathOU)
----------------------------------------
"@
    }elseif($assunto -like "*manual*"){
       $logFilePath = "C:\Util\Mendes - Scripts\Criacao de Usuario\Logs\UsuariosParaCriarManual.txt"
       $logContent = @"
----------------------------------------
Usuário deve ser criado!
    Nome do Colaborador: $($UserObject.nomeFormatado)
    Login: $($UserObject.nickname)
    Data de Admissão: $($UserObject.datAdm)
    EmployeeId: $($UserObject.employeeID)
    Senha: $($UserObject.password)
    Gestor E-mail: $($UserObject.emailGestor)
    OU: $($UserObject.pathOU)
----------------------------------------
"@
    }
    Add-Content -Path $logFilePath -Value $logContent

    }
function NewUsers($UserObject){
       
    New-ADUser -Name $UserObject.nomeFormatado `
    -SamAccountName $UserObject.nickname `
    -UserPrincipalName $UserObject.upn `
    -DisplayName $UserObject.nomeFormatado `
    -GivenName $UserObject.firstName `
    -Surname $UserObject.surname `
    -Enabled $true -ChangePasswordAtLogon $false `
    -AccountPassword (ConvertTo-SecureString $UserObject.password -AsPlainText -Force) `
    -Path $UserObject.pathOU `
    -Description "Usuário do Exchange"`
    -Office $UserObject.office`
    -EmailAddress $UserObject.upn
    
    Write-Host "Criando usuário..." -ForegroundColor Cyan
    Start-Sleep -Seconds 5

    $novoUsuarioPropriedades = Get-ADUser -Identity $UserObject.nickname
    Set-ADUser -Identity $UserObject.nickname -EmployeeID $UserObject.employeeID`
                -Title $UserObject.jobTitle `
                -Department $UserObject.department`
                -Company $UserObject.empresa

    Set-ADUser -Identity $UserObject.nickname -Add @{Pager="$($UserObject.pager)";extensionAttribute1=("$($UserObject.datAdm)");extensionAttribute2="$($UserObject.datNas)";extensionAttribute3=("$($UserObject.CPF)")}


    foreach ($group in $UserObject.memberOf) {
        $grupoExiste = Get-ADGroup -Filter { Name -eq $group } -ErrorAction SilentlyContinue
        if($grupoExiste){
            Add-ADGroupMember -Identity $group -Members $UserObject.nickname
        }
        else{
            Log-Error "Grupo não encontrado: $group - $($UserObject.nickname)"
        }

    }
    if($UserObject.gestorOu){Set-ADUser -Identity $UserObject.nickname -Manager $UserObject.gestorOu}

    Enable-ADAccount -Identity $UserObject.nickname

    return $true

}
function Output($user){
  return @"
   ----------------------------------------------------------
    Nome Completo    : $($user.nomeFormatado)
    Nickname         : $($user.nickname)
    Login            : $($user.upn)
    Senha            : $($user.password)
    employeeID       : $($user.employeeID)
    Data de Adm      : $($user.datAdm)
    Data de Nasc     : $($user.datNas)
    pathOU           : $($user.pathOU)
    Sigla            : $($user.pager)
    Cargo            : $($user.jobTitle)
    Departamento     : $($user.department)
    Empresa          : $($user.empresa)
    Email Gestor     : $($user.emailGestor)
    Employee Gestor  : $($user.employeeIDGestor)
    Nome do Gestor   : $($user.gestorName)
    ----------------------------------------------------------
"@
}
function Log-Error([String]$message) {
    $logFilePath = "C:\Util\Mendes - Scripts\Criacao de Usuario\Logs\Errors.log"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - ERROR: $message"
    Add-Content -Path $logFilePath -Value $logMessage
    }

if($PSDefaultParameterValues.Values -ne "srvdf145.dufrio.local"){$PSDefaultParameterValues.Add("*-AD*:Server","srvdf145.dufrio.local")}
$outputUsers =""
[String]$PathXMLUsersAdmitidos = "C:\Util\Mendes - Scripts\Criacao de Usuario\Logs\Usuários admitidos.xml"
$Users = Import-Clixml -Path $PathXMLUsersAdmitidos
$criarManualXML = @()

$Users | ForEach-Object{
$user = [Ajustes]::New($_)
if($user.nickname){
    $outputUsers += Output $user
    $obj = [CriarUsuario]::New($user)
} else {
    $criarManualXML += $user
}
}
$criarManualXML | Export-Clixml -Path "C:\Util\Mendes - Scripts\Criacao de Usuario\Logs\CriarUsuariosManual.xml"
return $outputUsers
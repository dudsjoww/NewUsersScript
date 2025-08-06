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
    $UserExistente = Get-ADUser -Filter {extensionAttributeCPF -eq $this.CPF}
    if(!$this.nickname -and !$UserExistente){
    EnviarEmail $User "manual" "email"
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
        "Senior",
        "SEC - Radius Users",
        "OTRS-Clientes"
    )
    if($departament -like "*VENDAS*"){
        $this.memberOf += ("IMP_$filialMemberOf-Vendas")
        $this.ramal=$true 
    }
    elseif($departament -like "*LOGISTICA*"){
        $this.memberOf += ("IMP_$filialMemberOf-Logistica")
    }
    elseif($departament -like "*SAC*"){
        $this.memberOf += ("IMP_$filialMemberOf-SAC", "Intelipost")
    }
    elseif($departament -like "*TI*"){
        $this.memberOf += ("IMP_$filialMemberOf-TI")
    }
    if($departament -like "*MARKETING*"){
        $this.memberOf += ("IMP_POA-MARKETING-COLOR")
    }
    if($departament -like "*RH*"){
        $this.memberOf += ("IMP_$filialMemberOf-RH")
    }
    if($departament -like "*ADMINISTRATIVO*"){
        $this.memberOf += ("IMP_$filialMemberOf-Administrativo")
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
        Write-Host "Sorting"
        for($i=0; $i -lt $nicknames.Length; $i++){
            $antName = ""
            #if($i -gt 0){
            $antName = $nicknames[$i - 1]
            #}
            $actualName = $nicknames[$i]
            Write-Host "nickname atual: $($nicknames[$i])"
            if(
                ($actualName -like "*$($nameParts[0])*" -and $antName -notlike "*$($nameParts[0])*" ) -and $antName -notlike "" -or 
                (($actualName -like "*$($nameParts[-1])*" -and $antName -notlike "*$($nameParts[0])*") -and $antName -notlike "" -and
                ($antName -notlike "*$($nameParts[-1])*"))
            ){
                Write-Host "Posicao $($i): $($actualName) -> $($antName)"
                $nicknames[$i - 1] = $actualName
                $nicknames[$i] = $antName
                $i = 0
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
            Write-Host $this.create
            if($this.create){
                Write-Host $UserObject
                CriarLog $UserObject "ajuste"

                EnviarEmail $UserObject "RHAJUSTE" "email"
                EnviarEmail $UserObject "gestorAJUSTE" $UserObject.emailGestor
                if($UserObject.ramal){CreateTicket $UserObject "RAMAL"}
                CreateTicket $UserObject "365"
                #CreateTicket $UserObject "kitAdm"
                CreateTicket $UserObject "kitInfra"
            }


        } else{
            Write-Host "Usuario Vai ser criado"
            $this.create = NewUsers $UserObject

            if($this.create){
                
                EnviarEmail $UserObject "RHCRIACAO" "email"
                EnviarEmail $UserObject "gestorCRIACAO" $UserObject.emailGestor
                CriarLog $UserObject "criacao"
                if($UserObject.ramal){CreateTicket $UserObject "RAMAL"}
                CreateTicket $UserObject "365"
                #CreateTicket $UserObject "kitAdm"
                CreateTicket $UserObject "kitInfra"
            }
        }
            #Quando criado/Modificado o usuario, sera enviado email correspondente para rh e gestor e log
            #Parametros: ObjetoUsuario para os campos do email, String para if e mudança do corpo, e para quem
        if(!$this.create){

            EnviarEmail $UserObject "manual" "email"
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
        $UserExistente = Get-ADUser -Filter {extensionAttributeCPF -eq $cpf} -ErrorAction SilentlyContinue
        if($UserExistente.UserPrincipalName -like "*$domain*" ){
            Write-Host "Usuario com mesmo dominio e CPF"
            return $true
    } else {
            Write-Host "Usuário nao encontrado" 
            return $false
        }
    }
    [bool]SetUser($UserObject){
        $UserExistente = Get-ADUser -Filter {extensionattributeCPF -eq $UserObject.CPF} -Properties SamAccountName, employeeid

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
        }
        return $true
    }
}
function EnviarEmail($UserObject, $assunto, $emailSend){
    [String]$to = $emailSend
    [String]$smtpServer = "smtp.sendgrid.net"
    [int]$smtpPort = " "
    [String]$username = ""
    [String]$password = "SENDGRID API KEY"
    [String]$from = "servicedesk@dufrio.com.br"
    if($assunto -like "*RHCRIACAO*"){
        [String]$subject = "Usuario Criado no AD - $($UserObject.nickname)"
        $body = "
            Usuario: $($UserObject.upn)
            EmployeeID: $($UserObject.employeeID)
            Nome: $($UserObject.nomeFormatado)
            Gestor: $(if($UserObject.emailGestor){$UserObject.emailGestor}else{"Sem gestor cadastrado"})
            Cargo: $($UserObject.jobTitle)
            Filial: $($UserObject.pager)
        "
    }
    if($assunto -like "*RHAJUSTE*"){
        $UserOld = Get-ADUser -Filter {extensionAttributeCPF -eq $UserObject.CPF}
        [String]$subject = "Usuario Alterado no AD - $($UserObject.SamAccountName)"
        $body = "
            Usuario: $($UserOld.UserPrincipalName)
            EmployeeID: $($UserObject.employeeid)
            Nome: $($UserOld.Name)
            Gestor: $(if($UserObject.emailGestor){$UserObject.emailGestor}else{"Sem gestor cadastrado"})
            Cargo: $($UserObject.jobTitle)
            Filial: $($UserObject.pager)
        "
    }
    elseif($assunto -like "*gestorAJUSTE*"){
        $UserOld = Get-ADUser -Filter {extensionAttributeCPF -eq $UserObject.CPF} 

        [String]$subject = "Usuario Modificado - $($UserOld.SamAccountName)"
        $body = "
        Se Usuario ja ativo, desconsidere a nova senha.

            Usuario: $($UserOld.UserPrincipalName)
            Senha: $($UserObject.password)
            Nome: $($UserOld.Name)
            Gestor: $(if($UserObject.emailGestor){$UserObject.emailGestor}else{"Sem gestor cadastrado"})
            Cargo: $($UserObject.jobTitle)
            Filial $($UserObject.pager)
        "
    
    }
     elseif($assunto -like "*gestorCRIACAO*"){
        [String]$subject = "Usuario Criado - $($UserObject.nomeFormatado)"
        $body = "
            Usuario: $($UserObject.upn)
            Senha: $($UserObject.password)
            Nome: $($UserObject.nomeFormatado)
            Gestor: $(if($UserObject.emailGestor){$UserObject.emailGestor}else{"Sem gestor cadastrado"})
            Cargo: $($UserObject.jobTitle)
            Filial $($UserObject.pager)
            
        "
    }
    elseif($assunto -like "*365*"){
        $UserOld = Get-ADUser -Filter {extensionAttributeCPF -eq $UserObject.CPF} 

        [String]$subject = "Criar usuario 365 - $($UserOld.SamAccountName)"
        $body = "
            Nao conseguiu criar chamado pelo OTRS, enviado via Email:
            Usuario: $($UserOld.UserPrincipalName)
            Nome: $($UserOld.Name)
            Gestor: $(if($UserObject.emailGestor){$UserObject.emailGestor}else{"Sem gestor cadastrado"})
            Cargo: $($UserObject.jobTitle)
            Data Admissao: $($UserObject.datAdm)
            Filial $($UserObject.pager)
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
    }elseif($assunto -like "*kitAdm*"){
        [String]$subject = "Reposicao de equipamentos - $($UserObject.pager)"
        $body = "
            Usuario: $($UserObject.upn)
            Nome: $($UserObject.nomeFormatado)
            Gestor: $($UserObject.emailGestor)
            Cargo: $($UserObject.jobTitle)
            Filial $($UserObject.pager)
        "
    }elseif($assunto -like "*kitInfra*"){
        [String]$subject = "Entrega de equipamentos ao colaborador novo - $($UserObject.nomeFormatado)"
        $body = "
            Usuario: $($UserObject.upn)
            Nome: $($UserObject.nomeFormatado)
            Gestor: $($UserObject.emailGestor)
            Cargo: $($UserObject.jobTitle)
            Filial $($UserObject.pager)
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
    $UserOld = Get-ADUser -Filter {extensionAttributeCPF -eq $UserObject.CPF}
       $logFilePath = "Logs\UsuariosCriadosCSenhaLogs.txt"
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
    }
    elseif($assunto -like "*manual*"){
       $logFilePath = "Logs\UsuariosParaCriarManual.txt"
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

    Set-ADUser -Identity $UserObject.nickname -Add @{Pager="$($UserObject.pager)";extensionAttribute1=("$($UserObject.datAdm)");extensionAttribute2="$($UserObject.datNas)";extensionAttributeCPF=("$($UserObject.CPF)")}


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
    $createdUser = Get-ADUser -Identity $UserObject.nickname
    if($createdUser){
        return $true
    } else{
        return $false
    }
    

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
    $logFilePath = "\Logs\Errors.log"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - ERROR: $message"
    Add-Content -Path $logFilePath -Value $logMessage
    }
function CreateTicket([PSCustomObject]$UserObject, [String]$subject){
#SESSION ID SERVICEDESK
	$header = @{
"Content-Type" = "application/json"
}

$url = ""

$body = @{
    UserLogin = ""
    Password  = ""
}

$jsonBody = $body | ConvertTo-Json -Compress

$response = Invoke-RestMethod -Uri $url -Method Post -Headers $header -Body $jsonBody

#createTicket



    $urlCreateTicket = ""
    $UserOld = Get-ADUser -Filter {extensionAttributeCPF -eq $UserObject.CPF}
    $bodyTicketMessage = @"
        Usuario: $($UserOld.UserPrincipalName)
        Nome: $($UserOld.Name)
        Gestor: $(if($UserObject.emailGestor){$UserObject.emailGestor}else{"Sem gestor cadastrado"})
        Cargo: $($UserObject.jobTitle)
        Data Admissao: $($UserObject.datAdm)
        Filial $($UserObject.pager)
"@ -replace "`r`n", "\n"
        
        $costumerUser = ""
        switch($subject){
        "365"{
            $queue = "Sistemas::N1"
            $title = "Criar/Alterar User 365"
            $costumerUser = "$($UserObject.emailGestor)"
            $serviceID = "2269"
            $slaID = "5" #ok
        }
        "kitAdm"{
                $queue = "Administrativo"
                $title = "Reposicao de Equipamentos - $($UserObject.pager)"
                $serviceID = "3480"
                $slaID = "11" #ok            
        }
        "kitInfra"{
            #if($UserObject.pager -match "FILIAIS"){
                $queue = "Infraestrutura::N1"
                $title = "Verificacao de usuario e Entrega de Equipamento $($UserObject.pager)"
                $serviceID = "2877"
                $slaID = "3" #ok
                $bodyTicketMessage = @"
                    Precisa validar a licenca do usuario e seus acessos:

                    Usuario: $($UserOld.UserPrincipalName)
                    Nome: $($UserOld.Name)
                    Gestor: $(if($UserObject.emailGestor){$UserObject.emailGestor}else{"Sem gestor cadastrado"})
                    Cargo: $($UserObject.jobTitle)
                    Data Admissao: $($UserObject.datAdm)
                    Filial $($UserObject.pager)
"@ -replace "`r`n", "\n"
            #}else{
                #$queue = "Administrativo"
               # $title = "Entrega de equipamento ADM $($UserObject.pager)"
               # $queue = "Administrativo::Controle de Ativos"
               # $serviceID = "2280"
               # $slaID = "17" #ok 
            #}
            
        }
        "RAMAL"{
            $queue = "Infraestrutura::N1"
            $title = "Criacao de Ramal"
            $serviceID = "1272"
            $slaID = "5" #ok
        }
        default{
            $queue = "Postmaster"
            $title = "Criacao de chamado - Usuario novo"
            $serviceID = "0"
            $slaID = "0" #ok
        }
    }


$urlCreateTicket = ""



$jsonBodyTicket = @"
{
  "SessionID": "$($response.SessionID)",
  "Ticket": {
    "Title": "$($title)",
    "Queue": "$($queue)",
    "lock": "unlock",
    "PriorityID": 4,
    "TypeID": 4,
    "ServiceID":$($serviceID),
    "SLAID": $($slaID),
    "State": "new",
    "CustomerUser": "$($costumerUser)",
    "owner": "otrs.devops"
  },
  "Article": {
    "Subject": "Script",
    "Body": "$($bodyTicketMessage)",
    "ContentType": "text/plain; charset=UTF-8",
    "communicationChannel": "Script",
    "SenderTypeID": 1
  }
}
"@
    

    try {
        # Tenta criar o ticket
        $responseCreatedTicket = Invoke-RestMethod -Uri $urlCreateTicket -Method Post -Headers $header -Body $jsonBodyTicket

        # Exibe a resposta (normalmente em caso de sucesso)
        
        
        $logFilePath = "C:\Util\Mendes - Scripts\Criacao de Usuario\Logs\UsuariosCriadosCSenhaLogs.txt"
        

        # Caso a resposta tenha um campo 'Error'
        if ($responseCreatedTicket.Error) {
            
             $textLog = "Erro retornado pela API: $($responseCreatedTicket.Error) $($subject)"
             Write-Host $textLog
             Log-Error $textLog
             EnviarEmail $UserObject $subject ""
        } else{
            $textLog = "Ticket criado com sucesso: $($responseCreatedTicket.TicketNumber) $($subject)"
            Write-Host $textLog
        }
        Add-Content -Path $logFilePath -Value $textLog
    }
    catch {
        # Tratamento de erro da requisição
        Write-Host "Falha ao criar o ticket."
        Write-Host "Mensagem de erro: $($_.Exception.Message)"
        EnviarEmail $UserObject $subject ""

        # Opcional: Exibir detalhes do erro completo
        if ($_.ErrorDetails) {
            Write-Host "Detalhes adicionais: $($_.ErrorDetails.Message)"
        }
    }

}

if($PSDefaultParameterValues.Values -ne ""){$PSDefaultParameterValues.Add("Get-ADUser:Server", "")}
$outputUsers =""
[String]$PathXMLUsersAdmitidos = "Logs\Usuários admitidos.xml"
$Users = Import-Clixml -Path $PathXMLUsersAdmitidos
$criarManualXML = @()

$Users | ForEach-Object{
    $user = [Ajustes]::New($_)
    if($user.nickname){
        $outputUsers += Output $user
        $obj = [CriarUsuario]::New($user)
    } 
    else {
        $criarManualXML += $user
    }
}

$criarManualXML | Export-Clixml -Path "\Logs\CriarUsuariosManual.xml"
return $outputUsers
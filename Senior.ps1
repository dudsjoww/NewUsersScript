class Senior{
[String]$url
[String]$urlGestor
[String]$caminhoArquivo
[hashtable]$headers
[DateTime]$dataAtual 
[String[]]$bodys
[String]$formattedDate
[String]$bodyGestor
[PSCustomObject[]]$Users
[String[]]$UsersSemGestores
$collection

Senior() {
	$this.url = "URL SENIOR"
	$this.urlGestor = "URL SENIOR GESTOR"
	$this.caminhoArquivo = "PATH DO ARQUIVO SENIOR\UsuariosAdmitidos.xml"
	$this.headers = @{
"Content-Type" = "text/xml"
}
	$this.dataAtual = (Get-Date).AddDays(-5)
	$this.formattedDate = $this.dataAtual.ToString("dd/MM/yyyy")
	$this.Users = @()
    $this.UsersSemGestores = @()
	$this.collection
	$this.bodys
	$this.bodys += @"
XML SENIOR
"@
	$this.bodys += @"
XML SENIOR
"@
	
}
[PSCustomObject] GetUserByEmployeeID($employeeID) {
try {
	# Procura o usuário no AD pelo EmployeeID
	$user = Get-ADUser -Filter { EmployeeID -eq $employeeID } -Properties EmployeeID
	return $user
} catch {
	# Se não encontrar o usuário, retorna $null
	return $null
}
}
[void]GetUsers(){
    $UsersSeniorRAW=@()
    if (Test-Path $this.caminhoArquivo) {
			    Remove-Item $this.caminhoArquivo -Force
		    }
    foreach($body in $this.bodys){
	    $response = Invoke-RestMethod -Uri $this.url -Method Post -Headers $this.headers -Body $body
	    if ($response) {
		    $this.collection = $response.Envelope.Body.ColaboradoresAdmitidosResponse.result.TMCSColaboradores
	    }
	    else {
		    #FAZER OUTRO IF ELSE EM CASO DE NÃO ENCONTRAR USERS!!!!!!!!!!!!!!!!!!!!!!!!
		    Write-Host "Nenhum conteúdo no Body encontrado."
	    }
	    if($this.collection){
		    $UsersSeniorRAW += $this.collection
		    $this.collection | ForEach-Object {
	        $employeeID = $_.numCad
	        # Verifica se o usuário existe no AD
	        $userInAD = $this.GetUserByEmployeeID($employeeID)
	        #Remove EmployeeIDs Ja Existentes
	        if ($userInAD) {
		        $this.collection = $this.collection | Where-Object { $_.numCad -ne $employeeID }
	        }
	        else {
		        #CRIAR FUNCAO PARA RETORNAR UM OBJETO
		        $NewUser = $this.CreateXMLUser($_)
                #ativar/comentar if ou desativar/descomentar de acordo com a necessidade do gestor
		        if($NewUser.employeeIDGestor){
		            $this.Users += $NewUser
		            $this.CriarLogs($NewUser)
		        }
                else{
                    #ENVIAR EMAIL/API DO SERVICEDESK <---------------------------
                }
                Write-Host "------------------------"
		        Write-Host $NewUser.name
		        Write-Host $employeeID
                if(!$NewUser.employeeIDGestor){ 
                    Write-Host "Sem gestor"
                    $this.UsersSemGestores += $employeeID    
                }
	        }
	    #$collection | Out-File "C:\Users\oliveira.eduardo\Desktop\Teste\Usuários TESTESADMITIDOS.txt"
    }
	
	    }


    #$this.collection | Out-File "C:\Users\oliveira.eduardo\Desktop\Teste\Criação de usuário 2.0\UsuariosAdmitidosRAWSENIOR.txt"

    #Cria log XML do que vai sair
    $this.Users | Export-Clixml -Path $this.caminhoArquivo
    
    $UsersSeniorRAW | Out-File "CAMINHO DE LOG.txt"
    Write-Host "Deus é fiel"
    }
    if($this.UsersSemGestores){
        $this.SendEmail($this.UsersSemGestores,"PREENCHER EMAIL AQUI")
    }

}
[PSCustomObject] GetGestorEmployeeID($employeeid){
$this.bodyGestor = @"
	BODY XML SENIOR
"@
$responseGestor = Invoke-RestMethod -Uri $this.urlGestor -Method Post -Headers $this.headers -Body $this.bodyGestor
return $responseGestor.Envelope.Body.ConsultarTabelasResponse.result.ocorrencia.resultado.campo.valor
 
}
[PSCustomObject]CreateXMLUser($user){

$employeeIDGestor = $this.GetGestorEmployeeID($user.numCad)
return [PSCustomObject]@{
	employeeID = $user.numCad
	employeeIDGestor = $employeeIDGestor
	posto = $user.codLoc
	nasc= $user.datNas
	admiss = $user.datAdm
	CPF = $user.numCpf
	name = $user.nomFun
	cargo = $user.titCar
} 
}
[void]CriarLogs($UserObject){
   $logFilePath = "LOG PATH\CriacaoUsuarioRAW.log"
   $logContent = @"
----------------------------------------
XML de Usuário criado:
Nome do Colaborador: $($UserObject.name)
EmployeeId: $($UserObject.employeeID)
EmployeeIdGestor: $($UserObject.employeeIDGestor)
Data de Admissão: $($UserObject.admiss)
Data de Nascimento: $($UserObject.nasc)
posto: $($UserObject.posto)
cargo: $($UserObject.cargo)
----------------------------------------
"@
Add-Content -Path $logFilePath -Value $logContent
}
    [void]SendEmail($employeeIDs, $emailSend){
    [String]$to = $emailSend
    [String]$smtpServer = "smtp.sendgrid.net"
    [int]$smtpPort = 
    [String]$username = ""
    [String]$password = ""
    #From é padrão
    [String]$from = ""
    
    $subject = "Atualizar tabelas Gestores"
    $bodyEmail = "
    Por favor, atualizar a tabela de gestores desses employeeIDs
    $($employeeIDs)

    "
    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($username, $securePassword)

    # Envie o e-mail
    Send-MailMessage -SmtpServer $smtpServer -Port $smtpPort -UseSsl -Credential $credential -From $from -To $to -Subject $subject -Body $bodyEmail
    }
}
$obj = [Senior]::new()
$obj.GetUsers()


$scriptPath = ".\Ajustes.ps1"
& $scriptPath
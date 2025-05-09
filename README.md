# PowerShell Script - Gerenciamento de Usuário no Active Directory

## 📋 Descrição

Este script PowerShell realiza o gerenciamento automatizado de usuários no Active Directory (AD). Ele executa as seguintes etapas:

1. Coleta o nome de um usuário.
2. Verifica se o usuário já existe no AD.
3. Se **não existir**:
   - Formata o nome de usuário.
   - Ajusta atributos necessários.
   - Adiciona o usuário aos grupos adequados.
   - Cria o usuário no AD.
4. Se **já existir**:
   - Move o usuário para uma unidade organizacional (OU) específica.
   - Se o usuário estiver **desabilitado**, o script:
     - Habilita a conta.
     - Reseta a senha.

## 🛠️ Requisitos

- PowerShell 5.1 ou superior
- Módulo `ActiveDirectory` (RSAT)
- Permissões administrativas no AD

## 🚀 Como usar

1. Abra o PowerShell como Administrador.
2. Execute o script:

```powershell
.\Gerenciar-UsuarioAD.ps1

//Para adicionar membros a Lista de Distribuição
Add-DistributionGroupMember –Identity “Nome do Grupo” –Member user@contoso.com

//Para dar permissão de Proprietário de uma lista de Distribuição
Set-DistributionGroup “Nome do Grupo” -ManagedBy “Admin@contoso.com” –BypassSecurityGroupManagerCheck

//Para remover um membro de uma Lista de Distribuição
Remove-DistributionGroupMember -Identity “Nome do Grupo” -Member user@contoso.com

//Para verificar membros de uma Lista de Distribuição
Get-DistributionGroupMember -identity “Nome do Grupo”|fl DisplayName,WindowsLiveID,RecipientType

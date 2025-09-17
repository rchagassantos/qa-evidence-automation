# PSR Automation - QA Evidence Standardization

Automação que transforma evidências do Gravador de Passos do Windows em documentos padronizados automaticamente, eliminando retrabalho no QA.

## Visão Geral

Fluxo:
Gravador de Passos → Salvar ZIP em C:\PSR_Automatico → WatchPSR detecta → Extrai → Padroniza → Salva DOCX em Desktop\Homologacao

## Como usar

1. Crie a pasta monitorada: `C:\PSR_Automatico`
2. Coloque os scripts `WatchPSR.vbs` e `PadronizarArraste_v3.vbs` dentro dela
3. Crie um atalho do `WatchPSR.vbs` na Área de Trabalho
4. Ao ligar ou reiniciar, clique duas vezes no atalho para iniciar o monitoramento
5. Use o Gravador de Passos (Step Recorder), grave os testes e salve o ZIP na pasta monitorada
6. A automação extrai, padroniza e salva o DOCX automaticamente

## Licença

Este projeto está licenciado sob a Apache License 2.0 - veja o arquivo LICENSE para mais detalhes.

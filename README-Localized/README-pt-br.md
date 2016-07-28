# Suplemento JavaScript Redact para Word

Saiba como criar um suplemento para pesquisar, realçar e editar texto.    

## Sumário
* [Histórico de Alterações](#change-history)
* [Observação importante](#important-note)
* [Pré-requisitos](#prerequisites)
* [Configurar o projeto](#configure-the-project)
* [Executar o projeto](#run-the-project)
* [Perguntas e comentários](#questions-and-comments)
* [Recursos adicionais](#additional-resources)

## Change History

25 de julho de 2016:
* Versão inicial do exemplo.

##Observação sobre importante

Neste exemplo, o texto é editado substituindo as letras por um asterisco.  O layout não é preservado como tal.  Não há realização de medidas ou cálculos para determinar a fonte, o tamanho ou a formatação das letras.

## Pré-requisitos

* Word 2016 para Windows, compilação 16.0.6912.1000 ou superior.
* [Node e npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) - Você deve usar uma versão mais recente já que as versões anteriores podem mostrar um erro ao gerar os certificados.

## Configurar o projeto

Execute os seguintes comandos do shell do Bash na raiz do projeto:

1. Clone este repositório em seu computador local.
2. ```npm install``` para instalar todas as dependências em package.json.
3. ```bash gen-cert.sh``` para criar os certificados necessários para executar este exemplo. 
* Em seguida, no repositório do computador local, clique duas vezes em ca.crt e escolha **Instalar Certificado**. 
* Escolha **Computador local** e **Avançar** para continuar. 
* Escolha **Colocar todos os certificados no repositório a seguir** e **Procurar**.  
* Escolha **Autoridades de Certificação Confiáveis** e selecione **OK**. 
* Escolha **Avançar** e **Concluir**; a partir daí, o certificado de autoridade de certificação é adicionado ao repositório de certificados.
4. ```npm start``` para iniciar o serviço de nó do servidor Web.

Agora, você precisa informar ao Microsoft Word onde encontrar o suplemento.

1. Crie um compartilhamento de rede ou [compartilhe uma pasta para a rede](https://technet.microsoft.com/pt-br/library/cc770880.aspx) e coloque o arquivo de manifesto [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml) nele.
3. Inicie o Word e abra um documento.
4. Escolha a guia **Arquivo** e escolha **Opções**.
5. Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.
6. Escolha **Catálogos de Suplementos Confiáveis**.
7. No campo **URL do Catálogo**, insira o caminho de rede para o compartilhamento da pasta que contém word-add-in-js-redact-manifest.xml e escolha **Adicionar Catálogo**.
8. Selecione a caixa de seleção **Mostrar no Menu** e, em seguida, escolha **OK**.
9. Será exibida uma mensagem para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Microsoft Office. Feche e reinicie o Word.

## Executar o projeto

1. Abra um documento do Word.
2. Na guia **Inserir** no Word 2016, escolha **Meus Suplementos**.
3. Selecione a guia **PASTA COMPARTILHADA**.
4. Escolha o **suplemento Word Redact** e marque **OK**.
5. Se os comandos de suplemento forem compatíveis com sua versão do Word, a interface do usuário informará que o suplemento foi carregado.

### Faixa de Opções da Interface do Usuário

Na Faixa de Opções:
* selecione a guia **Revisão** e escolha **Mostrar Painel de Tarefas de Redação** para inicializar o painel de tarefas.

 > Observação: O suplemento será carregado no painel de tarefas se os comandos de suplemento não forem compatíveis com sua versão do Word.

### Interface do usuário do painel de tarefas

No painel de tarefas, você pode:
* Pesquisar e realçar texto no documento, inserindo termos na caixa de pesquisa e selecionando o botão **Pesquisar e realçar**.
  
> OBSERVAÇÃO:  o recurso Pesquisar diferencia maiúsculas de minúsculas.  Para desfazer uma ação, pressione Ctrl+Z no teclado.

* Para editar texto no documento, digite termos na caixa de pesquisa e escolha o botão **Redact**.
  
> OBSERVAÇÃO:  o recurso Redact diferencia maiúsculas de minúsculas.   

* Para editar o texto selecionado no documento, escolha o botão **Redigir texto selecionado**.
  
> OBSERVAÇÃO:  o recurso Redact diferencia maiúsculas de minúsculas.       
  
### Comandos de suplemento

Este exemplo mostra como criar comandos de suplemento na Faixa de Opções. As declarações de comando do suplemento estão localizadas em word-add-in-js-redact-manifest.xml. 

## Perguntas e comentários

Gostaríamos de receber seus comentários sobre o exemplo do Word Redact. Você pode enviá-los na seção *Problemas* deste repositório.

As perguntas sobre o desenvolvimento do Office 365 em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Não deixe de marcar as perguntas com [office-js] e [API].

## Recursos adicionais

* [Documentação dos suplementos do Office](https://msdn.microsoft.com/pt-br/library/office/jj220060.aspx)
* [Centro de Desenvolvimento do Office](http://dev.office.com/)
* [Exemplos de código e projetos iniciais de APIs do Office 365](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## Direitos autorais
Copyright (C) 2016 Microsoft Corporation. Todos os direitos reservados.



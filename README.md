<h1>Bot para WhatsApp</h1>
Projeto em desenvolvimento.

<h2>Sobre</h2>
<p>Fiz esse bot para enviar mensagens a clientes cujo a data esteja próxima do vencimento com um dia de antecedência.<br>
A mensagem que é enviada contém o nome do cliente e sua data de vencimento, cada cliente recebe a mensagem com seu nome e sua data de vencimento, sendo possível a troca da mensagem e data de antecedência da forma que preferir.<br>

<h2>Como Funciona:</h2>
<p>O script foi feito para ser usado com planilhas. Logo é necessario ter uma planilha com os dados dos clientes, contendo as informações nessa ordem: <br> <br> Nome, Telefone, Dia do vencimento.<br>

É necessario também que a planilha esteja com o nome "Planilha Atualizada" para o funcionamento.<br> 

Caso você tenha uma planilha pronta, fora dessa ordem e com o nome diferente, o script tem uma funcionalidade que cria uma cópia dos dados dos clientes na ordem correta.<br>

Para isso é necessárioa colocar a planilha que deseja fazer a cópia na pasta "Planilha". Vale lembra que provavelmente sua planilha esteja com a posição diferente requerida no codigo, sendo necessário o ajuste do mesmo.<br> 

A pasta "planilha" é criada automaticamente após usar a opção "Criar lista de clientes".

Feito isso, criará uma planilha com o nome "Planilha Atualizada".<br> 

Agora está tudo pronto para iniciar o BOT.</p>

<h2>Funcionalidades:</h2>
<ul>
  <li> O bot abrirá o site para caso o celular não esteja conectado no WhatsApp Web para fazer essa configuração, após 20 segundos a pagina será fechada e o bot irá inciar o envio das mensagens </li><br>

  <li> O Bot tem um função que verifica se algum cliente está sem o número de telefone, alertando e registrando-o na "Planilha de Reenvio" com as infomções do cliente na pasta "Não Enviados" que será criada logo após o ocorrido. </li><br>

  <li> O Bot também tem uma simples verificação caso o número seja possivelmente inválido. Vale mencionar que o mesmo não garante que seja realmente um número inválido,
  caso ocorra algo que o impeça de enviar a mensagem será acionado o mesmo problema de número inválido. Com isso será também adicionado a planilha "Planilha de Reenvio" nas pasta "Não Enviados". </li><br>

  <li> Com a lista de clientes que não receberam a menssagem terá um controle para a verificação do cliente e assim poder corrigi-lo.</li>
</ul>


<h2>Seção de Pendências e Melhorias</h2>
<ul>
 <li> (em desenvolvimento) Criar funcionalidade para reenvio automática de mensagens não enviadas.
</ul>

<h2>Tecnologias</h2>
<div>
  <img src="https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54">
</div>

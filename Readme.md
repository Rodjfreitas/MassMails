# Mass&Mails

Mass&Mails é um aplicativo desenvolvido em python, utilizando as bibliotecas openpyxl, PySimpleGUI.
O aplicativo foi idealizado por mim, decorrente da necessidade encontrada no ambiente atual de trabalho, onde havia uma demanda excessiva e repetitiva de envios de e-mails de faturamento, onde o conteúdo do e-mail era sempre o mesmo, alterando-se apenas os anexos e os destinatários.

Já havia o hábito de salvar arquivos de extensão '.pdf' em pastas de nomes de clientes.
Com isso, desenvolvi inicialmente um código que recebia através de inputs, os diretórios do arquivo de excel onde eram informadas as listas de emails, nomes e diretórios da relação que precisava ser encaminhada.

Após uma configuração de login, e do conteúdo de e-mail, tal como assunto, com cópia, conteúdo e assinatura, o programa disparava os emails de acordo com cada linha do arquivo de excel, buscando no diretório informado, as pastas com nomes do clientes, e anexando o conteúdo '.pdf' inserido dentro das pastas. Após o envio, o programa emitia um relatório, com status de envio da relação.

A versão inicial, realizava envio de e-mails com e sem anexo, e não enviava os e-mails para as linhas onde o endereço de e-mail do destinatário principal não foi informado corretamente. Todas essas informações eram inseridas no relatório, que permitia análise posterior sobre os envios.

Após algumas análises, decidiu-se que os e-mails onde não haviam anexos, também não seriam enviados, uma vez que a função dos envios era enviar notas, não poderia enviar e-mails sem anexos pois faria com que o cliente recebesse mais de um e-mail, pois, após identificação do motivo do não envio, seria reenviado novamente com a correção.

Os principais motivos de envios de e-mails sem anexos, se davam:

###         1. Nomeclatura da pasta diferente da nomeclatura no arquivo
###         2. Sem arquivos salvos dentro da pasta
<br><br>


## Ferramentas Utilizadas:
### Python
### Dicionários em python
### Listas em python
### Condicionantes
### Repetições While, For, For com in enumerate()
### Funções
### Biblioteca Openpyxl
### Biblioteca PySimpleGUI
### Biblioteca smtplib
### Biblioteca email

<br><br>


## Versão 1.0 do Mass&Mails

Abaixo o menu principal da primeira versão:

![image](https://user-images.githubusercontent.com/119018022/232162858-5d88d1a4-3a7d-4bc7-8b02-004249af6c2e.png)

Abaixo o formato de dados necessários para inserção na planilha:

![Sem título](https://user-images.githubusercontent.com/119018022/232163381-5ccb42a4-74e7-4910-b967-944abb385980.png)


Esta versão só trabalha com servidor Gmail.

Nesta versão, na opção 5 (Informar Diretórios), o usuário inseria em primeiro lugar, o caminho do arquivo em excel onde estavam as informações de envio, e em segundo lugar, o local onde se encontravam as pastas com os arquivos. 

![image](https://user-images.githubusercontent.com/119018022/232163747-6a3e03f5-4027-416b-b734-1671098d9c0e.png)

Após essa inserção, ja era possível verificar pelo sistema, através da opção 3 (Lista de envios à realizar), a relação importada para envio.

![Sem título](https://user-images.githubusercontent.com/119018022/232164278-76f66a4b-0bed-49f3-9a51-a4f8f0c59cb8.png)

Antes de concluir o envio, realizava-se então a configuração das informações de email, tais como login, senha (não havia dispositivo de esconder senha), assunto, com cópia, conteúdo e assinatura.

![image](https://user-images.githubusercontent.com/119018022/232164834-b27dc37d-7690-4e7d-a623-f0e3f6ff1dea.png)


Abaixo o status após o envio. O usuário acompanha o processo de envio por mensagens no console.

![image](https://user-images.githubusercontent.com/119018022/232165025-d7aac87a-9d93-42a8-870d-b91bc5478ebd.png)


Para título de informação, observe que o e-mail da Natália não foi enviado, justamente porque há incosistência do nome na pasta e no arquivo (um possuí acentuação, outro não), e portanto o programa considera que não há pastas, não havendo pastas não há arquivos e portanto o e-mail não será enviado. 


Ao clicar na opção 5 (Impressão de lista de e-mails enviados), o usuário tem a opção de imprimir um relatório com o status de envio.


![image](https://user-images.githubusercontent.com/119018022/232165287-7a4a7685-e45e-4b0b-92f6-f19f1da2795e.png)


![image](https://user-images.githubusercontent.com/119018022/232165547-1e8f9cb1-518e-433a-998b-1c3675ad67e2.png)


## Caixa de E-mails

![image](https://user-images.githubusercontent.com/119018022/232168093-709fbb86-fa06-476c-bae9-a24f1d43cf81.png)


![image](https://user-images.githubusercontent.com/119018022/232168205-2c550cb8-c217-4e07-aa65-93372a2c1c74.png)


![image](https://user-images.githubusercontent.com/119018022/232167908-9426ff38-a88c-43f4-9d88-078e0f4b1ba9.png)



## Versão 2.0 do Mass&Mails

Buscando a evolução contínua, procurei informações para aplicar uma interface gráfica ao programa, para facilitar ainda mais nas ações do usuário.
Inicialmente havia me deparado com o Tkinter, uma biblioteca com essa finalidade, mas que por falta de tempo para exploração e conhecimento, inicialmente foi descartada. 

Me deparei então com a biblioteca PySimpleGUI. Simplesmente Fantástica. 
Intuitiva, de facil aprendizado e manuseio, me familiarizei desde o princípio. 

Eis o resultado obtido através dela (criei dois temas distintos, apenas para testes):

### 1. Tema Principal (meu favorito)

![image](https://user-images.githubusercontent.com/119018022/232166033-0e5ac5e0-f35b-4647-a4cd-1970cf6ba39e.png)


### 2. Tema DarkRed

![image](https://user-images.githubusercontent.com/119018022/232166004-10468736-4c1c-4a99-92d9-800a5ea2accd.png)


Diferente da v1.0, esta versão eu acrescentei a opção de escolha de servidor de e-mail. A primeira versão era habilitada apenas para Gmail, esta segunda, inclui no código as opções de escolhas para outlook e yahoo, outros dois serviços bastante utilizados atualmente.

O processo com a interface gráfica se apresnta bem mais simples, com o usuário podendo inserir as informações de forma mais rápida e intuitiva para encaminhamento de e-mail. Neste caso, houve mudança da planilha, uma vez que o assunto do e-mail não é mais inserido na planilha, e sim no programa.

![image](https://user-images.githubusercontent.com/119018022/232166606-114f709a-12a0-4a29-9f58-cdeac69639fc.png)


Ao clicar no botão "Browse", abre-se uma tela para o usuário escolher o local do arquivo e o diretório das pastas, respectivamente.


![image](https://user-images.githubusercontent.com/119018022/232166688-6e87778a-6506-4847-b498-2b9d2ba7293a.png)


Processo de Envio:


![image](https://user-images.githubusercontent.com/119018022/232167284-5df9be68-87c5-4c38-8d81-393b6e36872e.png)


Após clicar em ok, é emitido um relatório que é salvo na pasta do programa, sobre os status de envios.


![image](https://user-images.githubusercontent.com/119018022/232167363-79ffc646-363a-452e-85b9-3de52d60850a.png)



## Caixa de E-mails


![image](https://user-images.githubusercontent.com/119018022/232168370-73907f4b-c266-4763-b6e5-3540ff7eb87b.png)


![image](https://user-images.githubusercontent.com/119018022/232168488-64c6ea31-dad8-45ac-8f9d-76d84d1d04a1.png)


![image](https://user-images.githubusercontent.com/119018022/232168611-cecc10bd-93d1-4720-b5f0-13fb848e214b.png)


![image](https://user-images.githubusercontent.com/119018022/232171028-83d977e9-995f-49d4-9ad9-a959743846f6.png)










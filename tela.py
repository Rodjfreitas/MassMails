import PySimpleGUI as sg
import dados
import random

qtd_contagem = 0

# cores = ['Black', 'BlueMono', 'BluePurple', 'BrightColors', 'BrownBlue', 'Dark', 'Dark2', 'DarkAmber', 'DarkBlack', 'DarkBlack1', 'DarkBlue', 'DarkBlue1', 'DarkBlue10', 'DarkBlue11', 'DarkBlue12', 'DarkBlue13', 'DarkBlue14', 'DarkBlue15', 'DarkBlue16', 'DarkBlue17', 'DarkBlue2', 'DarkBlue3', 'DarkBlue4', 'DarkBlue5', 'DarkBlue6', 'DarkBlue7', 'DarkBlue8', 'DarkBlue9', 'DarkBrown', 'DarkBrown1', 'DarkBrown2', 'DarkBrown3', 'DarkBrown4', 'DarkBrown5', 'DarkBrown6', 'DarkBrown7', 'DarkGreen', 'DarkGreen1', 'DarkGreen2', 'DarkGreen3', 'DarkGreen4', 'DarkGreen5', 'DarkGreen6', 'DarkGreen7', 'DarkGrey', 'DarkGrey1', 'DarkGrey10', 'DarkGrey11', 'DarkGrey12', 'DarkGrey13', 'DarkGrey14', 'DarkGrey15', 'DarkGrey2', 'DarkGrey3', 'DarkGrey4', 'DarkGrey5', 'DarkGrey6', 'DarkGrey7', 'DarkGrey8', 'DarkGrey9', 'DarkPurple', 'DarkPurple1', 'DarkPurple2', 'DarkPurple3', 'DarkPurple4', 'DarkPurple5', 'DarkPurple6', 'DarkPurple7', 'DarkRed', 'DarkRed1', 'DarkRed2', 'DarkTanBlue', 'DarkTeal', 'DarkTeal1', 'DarkTeal10', 'DarkTeal11', 'DarkTeal12', 'DarkTeal2', 'DarkTeal3', 'DarkTeal4', 'DarkTeal5', 'DarkTeal6', 'DarkTeal7', 'DarkTeal8', 'DarkTeal9', 'Default', 'Default1', 'DefaultNoMoreNagging', 'GrayGrayGray', 'Green', 'GreenMono', 'GreenTan', 'HotDogStand', 'Kayak', 'LightBlue', 'LightBlue1', 'LightBlue2', 'LightBlue3', 'LightBlue4', 'LightBlue5', 'LightBlue6', 'LightBlue7', 'LightBrown', 'LightBrown1', 'LightBrown10', 'LightBrown11', 'LightBrown12', 'LightBrown13', 'LightBrown2', 'LightBrown3', 'LightBrown4', 'LightBrown5', 'LightBrown6', 'LightBrown7', 'LightBrown8', 'LightBrown9', 'LightGray1', 'LightGreen', 'LightGreen1', 'LightGreen10', 'LightGreen2', 'LightGreen3', 'LightGreen4', 'LightGreen5', 'LightGreen6', 'LightGreen7', 'LightGreen8', 'LightGreen9', 'LightGrey', 'LightGrey1', 'LightGrey2', 'LightGrey3', 'LightGrey4', 'LightGrey5', 'LightGrey6', 'LightPurple', 'LightTeal', 'LightYellow', 'Material1', 'Material2', 'NeutralBlue', 'Purple', 'Python', 'PythonPlus', 'Reddit', 'Reds', 'SandyBeach', 'SystemDefault', 'SystemDefault1', 'SystemDefaultForReal', 'Tan', 'TanBlue', 'TealMono', 'Topanga']
# cor_aleatoria = random.choice(cores)

def janelaPrincipal():
    # sg.change_look_and_feel('DarkEmber')
    sg.theme('DarkGrey14')
    input_size = (54, None)
    text_size = (20, None)
    # Layout
    layout = [
        # Selecionar o Arquivo
        [sg.Text('Selecionar Arquivo**', size=text_size, justification='right'), sg.FileBrowse('Pesquisar',
                target='arquivo'), sg.InputText(key='arquivo', size=(43, None))],
        # Selecionar a Pasta
        [sg.Text('Selecionar Diretório**', size=text_size, justification='right'), sg.FolderBrowse('Pesquisar',
                target='pasta'), sg.InputText(key='pasta', size=(43, None))],
        # Provedor
        [sg.Text('Provedor de E-mail**', size=text_size, justification='right'),
             sg.Radio('Gmail', 'provedor', key='gmail'), sg.Radio('Outlook', 'provedor', key='outlook'), sg.Radio('Yahoo', 'provedor', key='yahoo')],
        # Login e Senha
        [sg.Text('Usuário**', size=text_size, justification='right'),
             sg.Input(size=input_size, key='usuario')],
        [sg.Text('Senha**', size=text_size, justification='right'), sg.Input(
                key='senha', password_char='*', size=input_size)],
        # Assunto Email
        [sg.Text('Assunto**', size=text_size, justification='right'),
             sg.Input(size=input_size, key='assunto')],
        [sg.Text('Com Cópia', size=text_size, justification='right'),
             sg.Input(size=input_size, key='cc')],
        # Conteúdo do Email
        [sg.Text('Conteúdo**', justification='left')],
        [sg.Multiline(size=(60, 10), font=('Helvetica', 12),
                          key='conteudo', border_width=2)],
        # Botão de Envio
        [sg.Button('Disparar E-mails', size=33),
             sg.Button('Itens à Enviar', size=33)],
        [sg.Text('** Preenchimento Obrigatório.')],
        [sg.Text('Desenvolvido por Rodrigo Freitas', size=65, justification='center')]]    
    # Janela
    return sg.Window(
            'Mass&Mails v3.0 - Disparador Massivo de E-mails', layout=layout, finalize=True)


def janelaprogresso():
    sg.theme('DarkGrey14')
    relacao_a_enviar = dados.buscarnoArquivo(local_arquivo)
    qtd_contagem = len(relacao_a_enviar)
    #Layout
    layout = [
         [sg.Text('Processando Envios de E-mails.', justification='center')],
         [sg.Text('Por favor, aguarde', justification='center')],
         [sg.ProgressBar(qtd_contagem, orientation='h', size=(35, 20), key='progressbar')]]
    # Janela
    return sg.Window('Progresso dos Envios', layout=layout, finalize=True) 

def janelaitensEnviar():
    sg.theme('DarkGrey14')
    #Layout
    layout = [
        [sg.Text('Relação de Envios')],
        [sg.Output(size=(90, 30))],
        [sg.Button('Voltar')]]
    return sg.Window('Mass&Mails v3.0 - Relação de E-mails à Enviar', layout=layout, finalize=True)

#Criar janelas
janela1, janela2, janela3 = janelaPrincipal(), None, None

# Criar loop de Eventos
while True:
    window, event, values = sg.read_all_windows()    
    # Quando a janela for fechada
    if window == janela1 and event == sg.WIN_CLOSED:
        break
    if window == janela2 and event == sg.WIN_CLOSED:
        janela1.un_hide()
        janela2.hide()
    # Quando ir para a próxima janela
    ## Disparar Emails
    if window == janela1 and event == 'Disparar E-mails':            
        local_arquivo = values['arquivo']
        local_pasta = values['pasta']
        usuario_email = values['usuario']
        senha_usuario_email = values['senha']
        assunto_email = values['assunto']
        com_copia = values['cc']
        conteudo_email = values['conteudo']
        servidor_gmail = values['gmail']
        servidor_outlook = values['outlook']
        servidor_yahoo = values['yahoo']
        # Verificar se existe campos não preenchidos
        if usuario_email == "" or senha_usuario_email == "" or assunto_email == "" or conteudo_email == "":
            sg.popup('ATENÇÃO', 'Preenchimento de todos os campos obrigatórios')
            continue
        # Abrir janela de progresso
        janela3 = janelaprogresso()
        # Verificar qual servidor será usado
        if servidor_gmail:
            servidor = 'smtp.gmail.com'
        elif servidor_outlook:
            servidor = 'smtp.outlook.com'
        elif servidor_yahoo:
            servidor = 'smtp.yahoo.com'
        # Cria um lista com dados importados da planilha
        email_enviados = []
        relacao_a_enviar = dados.buscarnoArquivo(local_arquivo)
        # Loop de envio de e-mails
        qtd_contagem = len(relacao_a_enviar)
        janela3['progressbar'].update(0)
        for pos, valor in enumerate(relacao_a_enviar):
            janela3['progressbar'].update(pos + 1)
            if pos == 0:
                continue
            nome = valor[0]
            assunto = f'{assunto_email} - {nome}'
            conteudo = f'{dados.saudacao()},\n\n{conteudo_email}'
            email = valor[1]
            anexo = f'{local_pasta}\\{nome}'
            retorno_envio = dados.enviarEmail(usuario_email, senha_usuario_email, assunto,
                                            email, com_copia, conteudo, anexo, nome, servidor)
            if retorno_envio == 'Erro':
                sg.popup('ATENÇÃO', 'Verificar dados do Login.')
                janela3.close()
                status = False
                break
            status = True
            email_enviados.append(retorno_envio)
            assunto = ''
        if status:
            # Salvar relatório de envios no diretório
            dados.salvanoArquivo(email_enviados)
            janela3.hide()
            sg.popup_ok('Sucesso', 'Processo de Envio Concluído.', 'Salvo um relatório na pasta do programa.')               
            continue
    # Consultar itens à enviar
    if window == janela1 and event == 'Itens à Enviar':
        janela1.hide()
        janela2 = janelaitensEnviar()
        if values['arquivo'] == '':
            print('ATENÇÃO!\nNão foi informado caminho de arquivo na tela principal.\nInforme um arquivo de origem para que seja apresentado a relação de envios.')
        else:
            relacao_a_enviar = dados.buscarnoArquivo(values['arquivo'])
            dados.imprimirLista(relacao_a_enviar)
    # botão voltar do itens à enviar
    if window == janela2 and event == 'Voltar':
        janela1.un_hide()
        janela2.hide()
    

    # Quando voltar para a janela anterior

# Lógica no que deve acontecer ao clicar nos botões

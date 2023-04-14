import PySimpleGUI as sg
import dados

qtd_contagem = 0


class TelaPython:
    def __init__(self):
        sg.change_look_and_feel('DarkBrown4')
        input_size = (54, None)
        text_size = (20, None)
        # Layout
        layout = [
            # Selecionar o Arquivo
            [sg.Text('Selecione o Arquivo', size=text_size, justification='right'), sg.FileBrowse(
                target='arquivo'), sg.InputText(key='arquivo')],
            # Selecionar a Pasta
            [sg.Text('Selecionar Diretório', size=text_size, justification='right'), sg.FolderBrowse(
                target='pasta'), sg.InputText(key='pasta')],
            # Provedor
            [sg.Text('Provedor de E-mail:', size=text_size, justification='right'),
             sg.Radio('Gmail', 'provedor', key='gmail'), sg.Radio('Outlook', 'provedor', key='outlook'), sg.Radio('Yahoo', 'provedor', key='yahoo')],
            # Login e Senha
            [sg.Text('Usuário', size=text_size, justification='right'),
             sg.Input(size=input_size, key='usuario')],
            [sg.Text('Senha', size=text_size, justification='right'), sg.Input(
                key='senha', password_char='*', size=input_size)],
            # Assunto Email
            [sg.Text('Assunto', size=text_size, justification='right'),
             sg.Input(size=input_size, key='assunto')],
            [sg.Text('Com Cópia', size=text_size, justification='right'),
             sg.Input(size=input_size, key='cc')],
            # Conteúdo do Email
            [sg.Multiline(size=(60, 10), font=('Helvetica', 12),
                          key='conteudo', border_width=2)],
            # Barra de Progresso
            [sg.ProgressBar(qtd_contagem, orientation='h', size=(
                20, 20), key='progressbar', bar_color='gray', visible=False)],
            # Botão de Envio
            [sg.Button('Disparar E-mails', size=33, key='disparar_emails'),
             sg.Button('Emitir Relatório', size=33, disabled=True)],
            # Tela de Impressão
            [sg.Output(size=(77, 5))],
            [sg.Text('Desenvolvido por Rodrigo Freitas', size=65, justification='center')]]
        # Janela
        self.janela = sg.Window(
            'Mass&Mails v2.0 - Disparador Massivo de E-mails').layout(layout)

    def Iniciar(self):
        while True:
            # Extrair os dados da Tela
            self.button, self.values = self.janela.Read()

            # Sair sem erro no botão x
            if self.button == sg.WIN_CLOSED:
                break

            local_arquivo = self.values['arquivo']
            local_pasta = self.values['pasta']
            usuario_email = self.values['usuario']
            senha_usuario_email = self.values['senha']
            assunto_email = self.values['assunto']
            com_copia = self.values['cc']
            conteudo_email = self.values['conteudo']
            servidor_gmail = self.values['gmail']
            servidor_outlook = self.values['outlook']
            servidor_yahoo = self.values['yahoo']

            if usuario_email == "" or senha_usuario_email == "" or assunto_email == "" or conteudo_email == "":
                sg.popup('Erro', 'Preenchimento de todos os campos obrigatórios')
                continue

            if servidor_gmail:
                servidor = 'smtp.gmail.com'
            elif servidor_outlook:
                servidor = 'smtp.outlook.com'
            elif servidor_yahoo:
                servidor = 'smtp.yahoo.com'

            # Cria um lista com dados importados da planilha
            email_enviados = []
            relacao_a_enviar = dados.buscarnoArquivo(local_arquivo)

            # Habilita a barra de progresso
            self.janela['progressbar'].update(0)
            for pos, valor in enumerate(relacao_a_enviar):
                event, values = self.janela.read(timeout=10)
                if pos == 0:
                    continue
                nome = valor[0]
                assunto = f'{assunto_email} - {nome}'
                conteudo = f'{dados.saudacao()},\n{conteudo_email}'
                email = valor[1]
                anexo = f'{local_pasta}\\{nome}'
                retorno_envio = dados.enviarEmail(usuario_email, senha_usuario_email, assunto,
                                                  email, com_copia, conteudo, anexo, nome, servidor)
                if retorno_envio == 'Erro':
                    sg.popup('Erro', 'Verificar dados do Login.')
                    status = False
                    break
                else:
                    status = True
                    email_enviados.append(retorno_envio)
                    assunto = ''
                    self.janela['progressbar'].update(pos + 1)
            self.janela['progressbar'].update(0)
            if status:
                sg.popup_ok('Sucesso', 'Processo de Envio Concluído.')
                # Salvar relatório de envios no diretório
                dados.salvanoArquivo(email_enviados)
                continue


tela = TelaPython()
tela.Iniciar()

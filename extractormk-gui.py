import os, re, csv, sys, requests, time, pathlib, locale, configparser, ast
import PySimpleGUI as sg

locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
config = configparser.ConfigParser(allow_no_value=True)
pathcwd = str(pathlib.Path.cwd())

try:
    with open('config.ini') as f:
        config.read_file(f)
except IOError:
    sg.Popup("Arquivo de configuração não encontrado!")
    sys.exit()

config_host_url = config["default"]["host_url"]
config_login_mk_user = config["login_mk"]["usuario"]
config_login_mk_password = config["login_mk"]["senha"]

config = {
    'host_url': config_host_url
    }

login_payload = {
    'sys' : 'MK0',
    'user': config_login_mk_user,
    'password': config_login_mk_password,
    'action' : 'logon'
}

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36'}

contrato_csv_filename = "contrato.csv"

contrato_csv_headers = ["nome_razaosocial", "nome_fantasia", "cpf", "cnpj",
                        "rg", "data_nascimento", "email", "celular", "fonefixo",
                        "fonefax", "telefones", "cep", "cidade", "bairro",
                        "logradouro", "numero", "complemento", "data_vencimento",
                        "login_sac", "plano_cliente", "valor_plano",
                        "valor_plano_com_desconto", "vel_upload", "vel_download",
                        "meses_contrato", "descricao", "mac_address", 
                        "mac_address_aux", "contrato"]

s = requests.Session()

def login(config, login_payload, headers):
    host = config['host_url']
    
    if host:
        logon_url = host + "/mk/logon.do"
        try:
            r = s.post(logon_url, data=login_payload, headers=headers)
            login_status = True
        except requests.exceptions.RequestException as e:
            status_msg = "- Erro de conexão."
            login_status = False

        if login_status:
            data_url = host + "/mk/form.jsp?sys=MK0&action=openform&formID=7836"
            r = s.get(data_url, headers=headers)
            data = r.text
            e = re.compile(r"interactionError")
            if e.search(data) is not None:
                status_msg = "- Erro de autenticação."
                login_status = False
            else:
                status_msg = "- Autenticado."
    else:
        status_msg = "- Endereço de host inválido."
        login_status = False
    
    print(status_msg)
    
    return login_status

def obter_id(nomes):
    # Essa função filtra o ID dos nomes dos clientes no seguinte formato:
    # xxxx - NOME DO CLIENTE
    # xxxx é o ID que será trabalhado, o que vier após "-" é descartado.
    # Linhas com símbolo "+"" no final são ignoradas.
    # #FILTRO# na linha pode ser usada para emular a busca na janela de faturas,
    # recurso útil para baixar PDF de algum contrato específico.

    client_id_list = []
    
    for i in nomes.splitlines():
        if not len(i) <= 5 and not i.endswith("+"):
            filtro = i.partition("#")[-1].rpartition("#")[0]
            client_id = i.split("-")
            client_id = [x.strip() for x in client_id if x.strip()]
            if client_id[0].isdigit():
                client_id_list.append([client_id[0], filtro])
        
    return client_id_list

def consulta_id(c_id):
    host = config['host_url']

    data_url = host + "/mk/openform.do?sys=MK0&action=openform&formID=8&align=0&mode=-1&goto=-1&filter=mk_pessoas.codpessoa={}&scrolling=no".format(c_id)

    r = s.get(data_url, headers=headers)
    data = r.text

    nome_razaosocial = re.search(r"d\.c_33\ .*\'(.*?)\'\)\;", data).group(1)
    nome_fantasia = re.search(r"d\.c_602250\ .*\'(.*?)\'\)\;", data).group(1)
    cpf = re.search(r"d\.c_34\ .*\'(.*?)\'\)\;", data).group(1)
    cnpj = re.search(r"d\.c_35\ .*\'(.*?)\'\)\;", data).group(1)
    rg = re.search(r"d\.c_36\ .*\'(.*?)\'\)\;", data).group(1)
    data_nascimento = re.search(r"d\.c_38\ .*\'(.*?)\'\)\;", data).group(1)
    email = re.search(r"d\.c_52\ .*\'(.*?)\'\)\;", data).group(1)
    celular = re.search(r"d\.c_54\ .*\'(.*?)\'\)\;", data).group(1)
    fonefixo = re.search(r"d\.c_55\ .*\'(.*?)\'\)\;", data).group(1)
    fonefax = re.search(r"d\.c_56\ .*\'(.*?)\'\)\;", data).group(1)
    cep = re.search(r"d\.c_41\ .*\'(.*?)\'\)\;", data).group(1)
    cidade = re.search(r"d\.c_43\ .*\'(.*?)\'\)\;", data).group(1)
    bairro = re.search(r"d\.c_47\ .*\'(.*?)\'\)\;", data).group(1)
    logradouro = re.search(r"d\.c_45\ .*\'(.*?)\'\)\;", data).group(1)
    numero = re.search(r"d\.c_75\ .*\'(.*?)\'\)\;", data).group(1)
    complemento = re.search(r"d\.c_59\ .*\ '(.*?)\'\)\;", data).group(1)
    login_sac = re.search(r"d\.c_567049\ .*\'(.*?)\'\)\;", data).group(1)
    
    data_vencimento = re.search(r"d\.c_141610\ .*\'(.*?)\'\)\;", data).group(1)
    plano_padrao = re.search(r"d\.c_141611\ .*\'(.*?)\'\)\;", data).group(1)
    
    telefones = ", ".join(filter(None, [celular, fonefixo, fonefax]))

    dados_contrato = [nome_razaosocial, nome_fantasia, cpf, cnpj,
                    rg, data_nascimento, email, celular, fonefixo,
                    fonefax, telefones, cep, cidade, bairro,
                    logradouro, numero, complemento,
                    data_vencimento, login_sac, plano_padrao]

    print("--> {}...".format(dados_contrato[0]))

    refresh_gui_sleep(1)

    return dados_contrato

def consulta_detalhes_plano(plano_cliente):
    # Consulta a conexão e planos cadastrados.

    host = config['host_url']

    data_url = host + \
        "/mk/navigate.do?sys=MK0&action=navigate&formID=14&componentID=-1&type=1&q="

    r = s.get(data_url, headers=headers)
    data = r.text

    regex = r"\[\'{}\', '(.*?)', '(.*?)', \"\<div align\=right\>(.*?)\<\/div\>\", \"\<div align\=right\>(.*?)\<\/div\>\"\]".format(plano_cliente)
    m = re.search(regex, data)

    if m:
        vel_upload = m.group(1) + "K"
        vel_download = m.group(2) + "K"
        meses_contrato = m.group(3)
        valor_plano = m.group(4)

        # Retorna entrada com 10% desconto no valor do plano.
        p = float(valor_plano.replace(",", "."))
        d = 0.1
        t = p - (p * d)
        valor_plano_sem_desconto = locale.currency(p)
        valor_plano_com_desconto = locale.currency(
            float(format(t, ".1f")))

        plano = [plano_cliente, valor_plano_sem_desconto,
                valor_plano_com_desconto, vel_upload, vel_download,
                meses_contrato]
    else:
        plano = False
    
    return plano

def consulta_plano_conexao(c_id, plano_padrao=None):
    host = config['host_url']

    data_url = host + \
        "/mk/openform.do?sys=MK0&action=openform&formID=17&align=0&mode=-1&goto=-1&filter=codcliente={}&scrolling=no".format(
            c_id)

    r = s.get(data_url, headers=headers)
    data = r.text
       
    plano_conexao = re.search(r"d\.c_230\ .*\'(.*?)\'\)\;", data).group(1)
    mac_address = re.search(r"d\.c_232\ .*\'(.*?)\'\)\;", data).group(1)
    descricao = re.search(r"d\.c_240\ .*\'(.*?)\'\)\;", data).group(1)
    mac_address_aux = re.search(r"d\.c_233\ .*\'(.*?)\'\)\;", data).group(1)
    contrato = re.search(r"d\.c_371\ .*\'(.*?)\'\)\;", data).group(1)
    
    plano_info = None

    if plano_padrao:
        plano_info = consulta_detalhes_plano(plano_padrao)
    else:
        plano_info = consulta_detalhes_plano(plano_conexao)

    if plano_info:
        consulta = plano_info + [descricao, mac_address, mac_address_aux, contrato]
        return consulta
    else:
        return False
    
def gerador(lista_id, pdf=False, contrato=False, **params):
    pdf_info_list = []
    contrato_info_list = []

    if pdf or contrato:
        if lista_id:
            print("--- Consultando...")
            for item in lista_id:
                c_id = item[0]
                print("\n--- ID: ", c_id)
                contrato_data = consulta_id(c_id)
                
                nome_razaosocial = contrato_data[0]
                plano_padrao = contrato_data[-1]
                
                if pdf:
                    filtro = item[1]
                    pdf_id = gerar_fatura_pdf(c_id, filtro)
                    if pdf_id is False:
                        print("-- Fatura do cliente não foi gerada!")
                    else:
                        pdf_info = [nome_razaosocial, pdf_id]
                        pdf_info_list.append(pdf_info)
                
                if contrato:
                    plano_conexao_info = consulta_plano_conexao(c_id, plano_padrao)
                    if plano_conexao_info is False:
                        print("-- Plano do cliente não foi definido!")
                    else:
                        contrato_info = contrato_data[:-1] + plano_conexao_info
                        contrato_info_list.append(contrato_info)
                
                refresh_gui_sleep(0)
            
            itens = len(pdf_info_list + contrato_info_list)

            if itens > 0:
                if gravar_dados(pdf_info_list,
                                contrato_info_list):
                    return True
            else:
                print("\n- Nenhum dado para gravar.")
        else:
            print("\n- Nenhuma ID fornecida, verifique a lista de nomes.")
    else:
        print("\n- Marque ao menos uma opção.")

def gravar_dados(pdf_info_list, contrato_info_list):
    print("\n--- Gravando dados...")

    if pdf_info_list:
        iterable = (x for x in pdf_info_list if x is not None)
        count = 0
        for pdf_data in iterable:
            nome_razaosocial = pdf_data[0]
            pdf_id = pdf_data[1]            
            if pdf_id is not False or None:
                    salvar_pdf(nome_razaosocial, pdf_id)
                    count += 1
        if count > 0:
            print(">>> PDFs: {}".format(count))

    if contrato_info_list:
        count = 0
        try:
            with open(contrato_csv_filename, "w", newline="") as f:
                f_wrt = csv.writer(f, delimiter=";")
                f_wrt.writerow(contrato_csv_headers)
                for c in contrato_info_list:
                    if c is not False:
                            if c is not None:
                                f_wrt.writerow(c)
                                count += 1
                if count > 0:
                    print(">>> Contratos: {}".format(count))
        except IOError as x:
            sg.Popup("ERRO.", x.errno, x.strerror)
            sg.Popup("Sem permissão - Arquivo em uso.\n Feche o Word/Excel.")
            print("- Erro na gravação, feche o Word/Excel e tente novamente!")
            return 0

    print("\n-- Gravação concluída.")

def gerar_fatura_pdf(c_id, filtro):
    host = config['host_url']

    post_payload_faturas = {
        'action': 'executeRule',
        'pType': '2',
        'ruleName': 'contas_faturas_entrar_aux',
        'sys': 'MK0',
        'formID': '464567976',
        'parentRID': '-1',
        'P_0': c_id,
        'P_1': 'TDS',
        'P_2': 'N',
        'P_3': 'S',
        'P_4': 'N',
        'P_5': 'N',
        'P_6': 'N',
        'P_7': filtro,
        'F_0_1363636235': c_id
    }

    data_url = host + "/mk/executeRule.do"

    r = s.post(data_url, data=post_payload_faturas, headers=headers)
    data = r.text

    matches = re.findall(r"\[(\d+)\,", data, re.DOTALL)

    post_payload_emitir = {
        'action': 'executeRule',
        'pType': '2',
        'ruleName': 'contas_emitir_enviar_cobranca',
        'sys': 'MK0',
        'formID': '464567976',
        'parentRID': '-1',
        'P_0': '0',
        'P_1': 'ArrayInstance({})'.format(matches),
        'P_2': 'S2',
        'P_3': '0',
        'P_4': '0',
        'F_0_1363636235': c_id
    }
 
    if matches is not None:
        print("-- Gerando PDF no servidor...")
        refresh_gui_sleep(0)
        r = s.post(data_url, data=post_payload_emitir, headers=headers)
        data = r.text
        m = re.search(r"/tmp/(.*).pdf", data)
        
        if m:
            pdf_id = m.group(1)
        else:
            pdf_id = False    

    return pdf_id

def salvar_pdf(nome_razaosocial, pdf_id):
    if pdf_id is not None:
        host = config['host_url']
        data_url = host + '/mk/tmp/' + pdf_id + '.pdf'
        
        fn = nome_razaosocial + '.pdf'
        r = s.get(data_url)

        path = pathlib.Path.cwd() / 'pdf' / time.strftime("%Y-%m-%d")
        pathlib.Path(path).mkdir(parents=True, exist_ok=True)
        filepath = path / fn
        
        with filepath.open('wb') as f:
            f.write(r.content)

def refresh_gui_sleep(s):
    window.Refresh()
    time.sleep(s)
       
sg.ChangeLookAndFeel('DarkBlue')

layout = [
    [sg.Text('Host:     '), sg.InputText(default_text=config['host_url'],
                                        size=(42, 1), do_not_clear=True)],
    [sg.Text('Usuário: '), sg.InputText(default_text=login_payload['user'],
                                        size=(25, 1), do_not_clear=True)],
    [sg.Text('Senha:   '), sg.InputText(default_text=login_payload['password'],
                                        password_char='*', size=(25, 1), do_not_clear=True)],
    [sg.Multiline(size=(50,15), do_not_clear=True)],
    [sg.Checkbox('Baixar PDF', default=False), sg.Checkbox('Contrato',
                                default=False), sg.Submit('Confirmar')],
    [sg.Output(size=(50,10))]]

window  = sg.Window('Extrator - MK 2.x').Layout(layout)

while True:
    event, values = window.Read(timeout=0)
    
    config['host_url'] = values[0]
    login_payload['user'] = values[1]
    login_payload['password'] = values[2]
    
    if event is None or event == 'Exit':
        break

    if event == 'Confirmar':
        lista_clientes = values[3]
        
        params = {'pdf': values[4],
                  'contrato': values[5]}
        
        if login(config, login_payload, headers):
                lista_id = obter_id(lista_clientes)
                gerador(lista_id, **params)               

window.Close()

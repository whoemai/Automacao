# Documentação do Código

Aqui está uma documentação para o código fornecido. Esta documentação descreve a finalidade do código, suas principais funções e como ele funciona.

## Objetivo do Código

O código tem o objetivo de automatizar a geração de informações sobre servidores em nuvem, especificamente na plataforma Azure. Ele cria uma planilha com detalhes de servidores, compara-a com uma planilha existente para encontrar novos servidores e atualiza a planilha com as informações mais recentes.

## Bibliotecas Utilizadas

O código utiliza as seguintes bibliotecas Python:

- `azure.identity`: Usada para autenticar com a Azure.
- `azure.mgmt.compute`: Usada para interagir com recursos de computação na Azure.
- `openpyxl`: Usada para trabalhar com arquivos Excel.
- `pandas`: Usada para manipular e analisar dados em formato tabular.
- `concurrent.futures`: Usada para paralelizar a obtenção de informações de servidores.

## Variáveis de Configuração

- `resource_group_to_ignore`: Uma lista de nomes de grupos de recursos que devem ser ignorados durante a coleta de informações.
- `output_excel_file`: O caminho do arquivo Excel de saída que conterá as informações dos servidores.
- `pd_base_file`: O caminho do arquivo Excel base que será usado para comparação.
- `new_info_excel_file`: O caminho do arquivo Excel que contém as novas informações a serem adicionadas.
- `output_file`: O caminho do arquivo Excel que será gerado com as informações atualizadas.

## Funções Principais

### `status_servidor(credential, subscription_id, resource_group_name, servidor)`

Esta função obtém o status de um servidor na Azure. Ela usa as credenciais fornecidas, o ID da assinatura, o nome do grupo de recursos e o nome do servidor para recuperar o status do servidor. O status é então retornado.

### `process_vm(credential, subscription_id, resource_group, servidor)`

Esta função processa as informações de um servidor na Azure. Ela chama a função `status_servidor` para obter o status do servidor e formata as informações relevantes, como tags, tamanhos de VM, vCPUs e zonas. As informações formatadas são retornadas como uma lista.

### `extract_vcpus(vm_size)`

Esta função extrai o número de vCPUs de um tamanho de VM. Ela divide o tamanho em partes e retorna o número de vCPUs, ou "N/A" se não for possível extrair.

### `encontrar_linhas_novas(pd_base_file, new_info_excel_file)`

Esta função compara os hostnames entre o arquivo Excel base e o arquivo Excel com as novas informações. Ela identifica quais hostnames são novos e os imprime na tela.

### `adicionar_linhas(pd_base_file, new_info_excel_file, output_file)`

Esta função adiciona as novas linhas de servidor do arquivo Excel com novas informações ao arquivo Excel base, se o usuário optar por fazê-lo. O usuário é solicitado a confirmar a ação.

## Fluxo Principal

O código principal começa criando uma planilha Excel vazia e escrevendo cabeçalhos nela. Em seguida, ele usa uma ThreadPoolExecutor para processar informações de servidores em paralelo para várias assinaturas e grupos de recursos na Azure. Isso é feito para melhorar a eficiência.

Depois que todas as informações são coletadas e formatadas, elas são escritas na planilha Excel de saída.

## Execução do Código

O código pode ser executado para:

1. Coletar informações de servidores na Azure.
2. Comparar as informações com uma planilha Excel base.
3. Adicionar as novas informações à planilha base, se o usuário concordar.

Certifique-se de configurar as variáveis de configuração adequadas antes de executar o código e seguir as instruções apresentadas durante a execução para adicionar ou não as novas informações à planilha base.

Após a execução, o arquivo Excel de saída será gerado com as informações atualizadas dos servidores.